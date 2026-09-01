import { useEffect, useMemo, useRef, useState } from 'react'
import {
  AlertCircle,
  Archive,
  Bell,
  CalendarClock,
  CheckCircle2,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  ClipboardCheck,
  Download,
  FileSpreadsheet,
  Filter,
  Flame,
  FolderOpen,
  Gauge,
  GitBranch,
  History,
  LineChart,
  PanelLeftClose,
  PanelLeftOpen,
  RefreshCw,
  ShieldCheck,
  SlidersHorizontal,
  Table2,
  Trash2,
  X,
  Zap,
} from 'lucide-react'
import { buildReportModel, mergeFilters, selectIssueDetailExportRows } from '@/calculations/report'
import { buildSampleBundle } from '@/data/sampleData'
import { CHANGELOG, type ChangelogEntry } from '@/data/changelog'
import { exportReportDeck } from '@/export/pptx'
import { exportSlidesPdf } from '@/export/pdf'
import {
  applyDataQualityEdits,
  auditDataQuality,
  dataQualityRoleLabel,
  type DataQualityFinding,
} from '@/services/dataQuality'
import { expandImportFiles, filesFromDrop, importSpreadsheet } from '@/services/fileImport'
import {
  buildUpdatedIssueWorkbook,
  downloadUpdatedIssueWorkbook,
  prepareIssueWorkbook,
  reconcileIssueRows,
  type EnrichedIssueField,
  type ManualIssueEnrichment,
  type ManualIssueWorkbookKind,
  type PreparedIssueWorkbook,
} from '@/services/manualIssuesUpdate'
import { clearLegacyConnectionData, loadFilters, saveFilters } from '@/services/storage'
import { checkForUpdate } from '@/services/updateChecker'
import { downloadUpdateFile } from '@/services/updateDownload'
import type {
  AgingBucket,
  ElectricalPoint,
  ImportedSheetFile,
  IssueDetailRow,
  KpiMetric,
  ReportFilters,
  SheetBundle,
  SheetRecord,
  SheetRole,
  UpdateInfo,
  WeeklyIssuePoint,
  WeldingPoint,
} from '@/types'
import { compactNumber, deltaLabel, percent } from '@/utils/format'
import { compareWorkWeeks, formatDate } from '@/utils/workWeeks'

const ROLE_CONFIG: Record<SheetRole, { label: string; icon: typeof AlertCircle; color: string }> = {
  bimIssues: { label: 'BIM Issues Log', icon: AlertCircle, color: '#38bdf8' },
  mechanical: { label: 'Mechanical / Process', icon: ClipboardCheck, color: '#14b8a6' },
  electrical: { label: 'Electrical Inspections', icon: Zap, color: '#f59e0b' },
  welding: { label: 'Welding Signoffs', icon: Flame, color: '#22c55e' },
}

function emptySheet(name: string): SheetRecord {
  return {
    id: name,
    name,
    rows: [],
  }
}

const EMPTY_BUNDLE: SheetBundle = {
  source: 'empty',
  sheets: {
    bimIssues: emptySheet('BIM Issues Log'),
    mechanical: emptySheet('Mechanical / Process Inspection Log'),
    electrical: emptySheet('Electrical Inspection Log'),
    welding: emptySheet('Welding Signoffs by Work Week'),
  },
}

const PREVIEW_BUNDLE = buildSampleBundle()
const ISSUE_ROWS_PER_EXPORT_SLIDE = 14
const ISSUE_ROWS_PER_APP_SLIDE = 12
const EXPORT_COOLDOWN_MS = 3000

interface DataQualityReview {
  bundle: SheetBundle
  imports: Partial<Record<SheetRole, ImportedSheetFile>>
  findings: DataQualityFinding[]
  values: Record<string, string>
}

function waitForExportCooldown(): Promise<void> {
  return new Promise((resolve) => window.setTimeout(resolve, EXPORT_COOLDOWN_MS))
}

function chunkRows<T>(rows: T[], size: number): T[][] {
  if (rows.length === 0) return [[]]
  return Array.from({ length: Math.ceil(rows.length / size) }, (_, index) => rows.slice(index * size, (index + 1) * size))
}

function cx(...classes: Array<string | false | null | undefined>): string {
  return classes.filter(Boolean).join(' ')
}

function Modal({
  open,
  title,
  children,
  onClose,
  wide,
  dismissible = true,
}: {
  open: boolean
  title: string
  children: React.ReactNode
  onClose: () => void
  wide?: boolean
  dismissible?: boolean
}) {
  if (!open) return null
  return (
    <div className="modal-backdrop" role="presentation" onMouseDown={dismissible ? onClose : undefined}>
      <div
        className={cx('modal', wide && 'modal-wide')}
        role="dialog"
        aria-modal="true"
        aria-label={title}
        onMouseDown={(event) => event.stopPropagation()}
      >
        <div className="modal-header">
          <h2>{title}</h2>
          {dismissible && (
            <button className="icon-button" type="button" onClick={onClose} aria-label="Close">
              <X size={18} />
            </button>
          )}
        </div>
        {children}
      </div>
    </div>
  )
}

function DataQualityModal({
  open,
  findings,
  values,
  onChange,
  onCancel,
  onApply,
}: {
  open: boolean
  findings: DataQualityFinding[]
  values: Record<string, string>
  onChange: (id: string, value: string) => void
  onCancel: () => void
  onApply: () => void
}) {
  const missingCount = findings.filter((finding) => !values[finding.id]?.trim()).length
  const suggestedCount = findings.filter((finding) => finding.autoFilled).length
  const affectedRows = new Set(findings.map((finding) => `${finding.role}:${finding.rowIndex}`)).size
  const orderedFindings = [...findings].sort((a, b) => {
    const aMissing = values[a.id]?.trim() ? 1 : 0
    const bMissing = values[b.id]?.trim() ? 1 : 0
    return aMissing - bMissing || a.role.localeCompare(b.role) || a.rowNumber - b.rowNumber || a.fieldLabel.localeCompare(b.fieldLabel)
  })

  return (
    <Modal open={open} title="Complete Missing Report Data" onClose={onCancel} wide dismissible={false}>
      <div className="data-quality-body">
        <div className="data-quality-intro">
          <span><AlertCircle size={20} /></span>
          <div>
            <strong>Review {findings.length} blank {findings.length === 1 ? 'field' : 'fields'} before generating the report</strong>
            <p>Contractor blanks are suggested as Bechtel. Every value can be changed before it is applied.</p>
          </div>
        </div>
        <div className="data-quality-summary" aria-label="Missing data summary">
          <span><strong>{affectedRows}</strong>Affected rows</span>
          <span className={cx(missingCount > 0 && 'needs-input')}><strong>{missingCount}</strong>Needs input</span>
          <span className="suggested"><strong>{suggestedCount}</strong>Suggested values</span>
        </div>
        <div className="data-quality-list">
          {orderedFindings.map((finding) => {
            const value = values[finding.id] ?? ''
            return (
              <label className={cx('data-quality-row', !value.trim() && 'missing')} key={finding.id}>
                <span className="data-quality-row-context">
                  <strong>{dataQualityRoleLabel(finding.role)}</strong>
                  <small>{finding.recordLabel} · Source row {finding.rowNumber}</small>
                </span>
                <span className="data-quality-field-label">
                  {finding.fieldLabel}
                  {finding.autoFilled && <em>Suggested</em>}
                </span>
                <input
                  aria-label={`${dataQualityRoleLabel(finding.role)} ${finding.recordLabel} ${finding.fieldLabel}`}
                  placeholder={`Enter ${finding.fieldLabel}`}
                  value={value}
                  onChange={(event) => onChange(finding.id, event.target.value)}
                />
              </label>
            )
          })}
        </div>
        <div className="data-quality-actions">
          <span>
            {missingCount > 0
              ? `${missingCount} required ${missingCount === 1 ? 'value remains' : 'values remain'}`
              : <><CheckCircle2 size={14} /> Ready to apply</>}
          </span>
          <button className="button secondary" type="button" onClick={onCancel}>Cancel import</button>
          <button className="button primary" type="button" disabled={missingCount > 0} onClick={onApply}>
            <CheckCircle2 size={16} />
            Apply corrections
          </button>
        </div>
      </div>
    </Modal>
  )
}

function UpdateModal({
  open,
  onClose,
  info,
  defaultTab,
  checking,
  lastChecked,
  onCheck,
}: {
  open: boolean
  onClose: () => void
  info: UpdateInfo | null
  defaultTab: 'update' | 'changelog'
  checking: boolean
  lastChecked: Date | null
  onCheck: () => Promise<void>
}) {
  const [tab, setTab] = useState<'update' | 'changelog'>(defaultTab)
  const [downloading, setDownloading] = useState(false)
  const [downloadedFilename, setDownloadedFilename] = useState<string | null>(null)
  const [downloadError, setDownloadError] = useState<string | null>(null)
  const [expandedVersions, setExpandedVersions] = useState<Set<string>>(
    () => new Set(CHANGELOG.length > 0 ? [CHANGELOG[0].version] : []),
  )
  const [showAll, setShowAll] = useState(false)

  useEffect(() => {
    if (open) {
      setTab(defaultTab)
      setDownloadedFilename(null)
      setDownloadError(null)
    }
  }, [open, info?.version, defaultTab])

  const changelogWithUpdate = useMemo(() => {
    if (!info?.version || CHANGELOG.some((entry) => entry.version === info.version)) return CHANGELOG
    const notes = info.releaseNotes
      .split('\n')
      .map((line) => line.trim().replace(/^[-*#]+\s*/, ''))
      .filter((line) => Boolean(line) && !/^what changed$/i.test(line) && !/^validation$/i.test(line))
    const synthetic: ChangelogEntry = {
      version: info.version,
      date: new Date().toISOString(),
      type: 'fix',
      notes: notes.length > 0 ? notes : ['A new packaged dashboard build is available.'],
    }
    return [synthetic, ...CHANGELOG]
  }, [info])

  useEffect(() => {
    const latest = changelogWithUpdate[0]
    if (!latest) return
    setExpandedVersions(new Set([latest.version]))
  }, [changelogWithUpdate])

  const visibleEntries = showAll ? changelogWithUpdate : changelogWithUpdate.slice(0, 7)
  const releaseNotes = info?.releaseNotes
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean) ?? []

  function toggleVersion(version: string): void {
    setExpandedVersions((previous) => {
      const next = new Set(previous)
      if (next.has(version)) next.delete(version)
      else next.add(version)
      return next
    })
  }

  function formatReleaseDate(value: string): string {
    const date = new Date(value)
    if (Number.isNaN(date.getTime())) return ''
    return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })
  }

  return (
    <Modal open={open} title={info ? 'Update Available' : 'Changelog'} onClose={onClose} wide>
      <div className="update-tabs">
        {info && (
          <button className={cx(tab === 'update' && 'active')} type="button" onClick={() => setTab('update')}>
            <Download size={14} />
            Update
          </button>
        )}
        <button className={cx(tab === 'changelog' && 'active')} type="button" onClick={() => setTab('changelog')}>
          <History size={14} />
          Changelog
        </button>
      </div>
      {tab === 'update' && info ? (
        downloadedFilename ? (
          <div className="update-success">
            <span className="update-success-icon"><CheckCircle2 size={30} /></span>
            <div>
              <strong>Update download started</strong>
              <p>{downloadedFilename} is being saved to your Downloads folder.</p>
            </div>
            <div className="update-next-step">
              Open the downloaded HTML for the new version. Your spreadsheet files remain separate from the app.
            </div>
            <button className="button secondary" type="button" onClick={onClose}>Done</button>
          </div>
        ) : (
          <div className="update-card">
            <div className="version-hop">
              <span>v{__APP_VERSION__}</span>
              <GitBranch size={15} />
              <strong>v{info.version}</strong>
            </div>
            <div className="release-notes">
              {(releaseNotes.length > 0 ? releaseNotes : ['A new packaged dashboard build is available.']).map((line, index) => {
                const heading = line.match(/^#{1,4}\s+(.+)/)
                if (heading) return <h4 key={`${line}-${index}`}>{heading[1]}</h4>
                return <p key={`${line}-${index}`}>{line.replace(/^[-*]\s+/, '')}</p>
              })}
            </div>
            <div className="update-download-note">
              <ShieldCheck size={17} />
              <div>
                <strong>Direct packaged download</strong>
                <span>The HTML is fetched inside this app and saved locally. No GitHub page will open.</span>
              </div>
            </div>
            {downloadError && <p className="download-error">{downloadError}</p>}
            <div className="modal-actions">
              <button className="button secondary" type="button" onClick={onClose}>Skip this version</button>
              <button
                className="button primary"
                type="button"
                disabled={downloading}
                onClick={async () => {
                  setDownloading(true)
                  setDownloadError(null)
                  try {
                    const result = await downloadUpdateFile(info)
                    setDownloadedFilename(result.savedFilename)
                  } catch {
                    setDownloadError('The packaged update could not be downloaded inside the app. Check your network or proxy and try again.')
                  } finally {
                    setDownloading(false)
                  }
                }}
              >
                {downloading ? <RefreshCw size={16} className="spin" /> : <Download size={16} />}
                {downloading ? 'Preparing update' : `Download v${info.version}`}
              </button>
            </div>
          </div>
        )
      ) : (
        <div className="changelog-panel">
          <div className={cx('update-status', info && 'has-update')}>
            <span className="update-status-icon">{info ? <Bell size={18} /> : <CheckCircle2 size={18} />}</span>
            <div>
              <strong>{info ? `v${info.version} is available` : `v${__APP_VERSION__} is up to date`}</strong>
              <span>{lastChecked ? `Last checked ${lastChecked.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' })}` : 'Checks automatically while the app is open'}</span>
            </div>
            <button className="button secondary compact" type="button" disabled={checking} onClick={() => void onCheck()}>
              <RefreshCw size={14} className={cx(checking && 'spin')} />
              {checking ? 'Checking' : 'Check now'}
            </button>
          </div>
          <div className="changelog-list">
            {visibleEntries.map((entry, index) => {
              const expanded = expandedVersions.has(entry.version)
              return (
                <article className={cx('changelog-entry', `type-${entry.type}`, expanded && 'expanded')} key={entry.version}>
                  <button type="button" className="changelog-entry-header" onClick={() => toggleVersion(entry.version)}>
                    <ChevronDown size={15} className={cx(!expanded && 'collapsed')} />
                    <strong>v{entry.version}</strong>
                    <span className="release-type">{entry.type}</span>
                    {index === 0 && <span className="latest-label">Latest</span>}
                    <time>{formatReleaseDate(entry.date)}</time>
                  </button>
                  {expanded && (
                    <ul>
                      {entry.notes.map((note) => <li key={note}>{note}</li>)}
                    </ul>
                  )}
                </article>
              )
            })}
          </div>
          {!showAll && changelogWithUpdate.length > visibleEntries.length && (
            <button className="show-all-releases" type="button" onClick={() => setShowAll(true)}>
              Show all {changelogWithUpdate.length} releases
            </button>
          )}
        </div>
      )}
    </Modal>
  )
}

function Sparkline({ data, hero }: { data: number[]; hero?: boolean }) {
  const w = 100
  const h = hero ? 34 : 22
  const pad = 3
  const min = Math.min(...data)
  const max = Math.max(...data)
  const span = max - min || 1
  const step = (w - pad * 2) / Math.max(1, data.length - 1)
  const points = data.map((value, index) => ({
    x: pad + index * step,
    y: pad + (h - pad * 2) * (1 - (value - min) / span),
  }))
  const line = points.map((point, index) => `${index === 0 ? 'M' : 'L'} ${point.x.toFixed(1)} ${point.y.toFixed(1)}`).join(' ')
  const last = points[points.length - 1]
  const area = `${line} L ${last.x.toFixed(1)} ${h - pad} L ${points[0].x.toFixed(1)} ${h - pad} Z`
  // The SVG is stretched non-uniformly, so the end marker is a real round HTML
  // element positioned over the line rather than a squashed SVG circle.
  return (
    <div className={cx('kpi-spark', hero && 'hero')} aria-hidden="true">
      <svg className="kpi-spark-svg" viewBox={`0 0 ${w} ${h}`} preserveAspectRatio="none">
        <path className="spark-area" d={area} />
        <path className="spark-line" d={line} vectorEffect="non-scaling-stroke" />
      </svg>
      <span className="kpi-spark-dot" style={{ left: `${last.x}%`, top: `${(last.y / h) * 100}%` }} />
    </div>
  )
}

function KpiTile({ metric, hero, hideDelta }: { metric: KpiMetric; hero?: boolean; hideDelta?: boolean }) {
  const Icon = metric.icon
  const toneClass = metric.tone === 'good' ? 'good' : metric.tone === 'warn' ? 'warn' : metric.tone === 'bad' ? 'bad' : ''
  const arrow = metric.delta > 0.0001 ? '▲' : metric.delta < -0.0001 ? '▼' : ''
  const deltaText = metric.deltaLabel
  return (
    <article className={cx('kpi-tile', hero && 'hero')}>
      <div className="kpi-tile-head">
        <span className="kpi-icon">
          <Icon size={hero ? 15 : 13} />
        </span>
        <span className="kpi-label">{metric.label}</span>
      </div>
      <div className="kpi-value-row">
        <strong>{metric.value}</strong>
        {!hideDelta && (
          <span className={cx('kpi-delta', toneClass)}>
            {arrow && <span aria-hidden="true">{arrow}</span>}
            {deltaText}
          </span>
        )}
      </div>
      {metric.spark && <Sparkline data={metric.spark} hero={hero} />}
    </article>
  )
}

const KPI_GROUPS = {
  primary: ['total-closed', 'total-opened', 'remaining-open', 'closure-rate'],
  week: ['opened-week', 'closed-week', 'inspections', 'sors'],
} as const

function KpiZone({ report }: { report: ReturnType<typeof buildReportModel> }) {
  const byId = new Map(report.kpis.map((metric) => [metric.id, metric]))
  return (
    <div className="kpi-groups">
      <section className="kpi-group this-week">
        <div className="kpi-group-head">
          <span>{report.periodLabel}</span>
        </div>
        <div className="kpi-group-body">
          {KPI_GROUPS.week.map((id) => {
            const metric = byId.get(id)
            return metric ? <KpiTile key={id} metric={metric} hideDelta /> : null
          })}
        </div>
      </section>
      <section className="kpi-group primary">
        <div className="kpi-group-head">
          <span>Project to date</span>
          <em>Through {report.periodEndWeek.label}</em>
        </div>
        <div className="kpi-group-body">
          {KPI_GROUPS.primary.map((id) => {
            const metric = byId.get(id)
            return metric ? <KpiTile key={id} metric={metric} /> : null
          })}
        </div>
      </section>
    </div>
  )
}

function FilterMenu({
  label,
  icon,
  options,
  selected,
  onChange,
  disabled,
}: {
  label: string
  icon: React.ReactNode
  options: string[]
  selected: string[]
  onChange: (next: string[]) => void
  disabled?: boolean
}) {
  const [open, setOpen] = useState(false)
  const menuRef = useRef<HTMLDivElement>(null)
  const active = !disabled && selected.length > 0
  const display = disabled
    ? 'Report range'
    : selected.length === 0
      ? 'All'
      : selected.length === 1
        ? selected[0]
        : 'Multiple'

  useEffect(() => {
    if (!open) return
    function handlePointerDown(event: PointerEvent): void {
      if (menuRef.current && !menuRef.current.contains(event.target as Node)) setOpen(false)
    }
    function handleKey(event: KeyboardEvent): void {
      if (event.key === 'Escape') setOpen(false)
    }
    document.addEventListener('pointerdown', handlePointerDown)
    document.addEventListener('keydown', handleKey)
    return () => {
      document.removeEventListener('pointerdown', handlePointerDown)
      document.removeEventListener('keydown', handleKey)
    }
  }, [open])

  return (
    <div className="filter-menu" ref={menuRef}>
      <button className={cx('filter-button', active && 'active')} type="button" onClick={() => setOpen((value) => !value)} disabled={disabled}>
        {icon}
        <span>{label}</span>
        <strong>{display}</strong>
        {selected.length > 1 && !disabled && <span className="count-badge">{selected.length}</span>}
        <ChevronDown size={14} />
      </button>
      {open && !disabled && (
        <div className="filter-popover">
          <button className="filter-clear" type="button" onClick={() => onChange([])}>
            All {label}
          </button>
          {options.map((option) => {
            const checked = selected.includes(option)
            return (
              <label className="filter-option" key={option}>
                <input
                  type="checkbox"
                  checked={checked}
                  onChange={() => {
                    if (checked) onChange(selected.filter((item) => item !== option))
                    else onChange([...selected, option])
                  }}
                />
                <span>{option}</span>
              </label>
            )
          })}
        </div>
      )}
    </div>
  )
}

type PlotPoint = {
  x: number
  y: number
}

// Round to a "nice" axis with clean integer-friendly ticks (never 37.5-style labels).
function niceScale(rawMax: number, targetSteps = 4): { max: number; ticks: number[] } {
  const safe = Math.max(rawMax, 1)
  const rough = safe / targetSteps
  const magnitude = 10 ** Math.floor(Math.log10(rough))
  const normalized = rough / magnitude
  const step = (normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 2.5 ? 2.5 : normalized <= 5 ? 5 : 10) * magnitude
  const max = Math.ceil(safe / step) * step
  const ticks: number[] = []
  for (let value = max; value >= -1e-9; value -= step) ticks.push(Math.round(value * 1000) / 1000)
  return { max, ticks }
}

function smoothPath(points: PlotPoint[]): string {
  if (points.length === 0) return ''
  return points.reduce<string>((path, point, index) => {
    if (index === 0) return `M ${point.x} ${point.y}`
    const previous = points[index - 1]
    const control = (point.x - previous.x) / 3
    return `${path} C ${previous.x + control} ${previous.y}, ${point.x - control} ${point.y}, ${point.x} ${point.y}`
  }, '')
}

function areaPath(points: PlotPoint[], baseline: number): string {
  if (points.length === 0) return ''
  return `${smoothPath(points)} L ${points[points.length - 1].x} ${baseline} L ${points[0].x} ${baseline} Z`
}

function chartValue(value: number): string {
  return Number.isInteger(value) ? String(value) : value.toFixed(1)
}

function workWeekNumber(label: string): string {
  const match = label.match(/^WW(\d{1,2})/i)
  return match ? `WW${match[1].padStart(2, '0')}` : label
}

function ChartTooltip({
  title,
  entries,
  xPercent = 12,
  compact,
}: {
  title: string
  entries: string[]
  xPercent?: number
  compact?: boolean
}) {
  return (
    <div
      className={cx('chart-tooltip', compact && 'compact')}
      style={{ left: `${Math.max(11, Math.min(89, xPercent))}%` }}
    >
      <strong>{title}</strong>
      {entries.map((entry) => <span key={entry}>{entry}</span>)}
    </div>
  )
}

function RangeChart({ data }: { data: WeeklyIssuePoint[] }) {
  const [hoveredIndex, setHoveredIndex] = useState<number | null>(null)
  const visible = data.slice(-30)
  const width = 980
  const height = 82
  const pad = { l: 40, r: 44, t: 10, b: 20 }
  const chartW = width - pad.l - pad.r
  const chartH = height - pad.t - pad.b
  const group = chartW / Math.max(visible.length, 1)
  const lineScale = niceScale(Math.max(1, ...visible.map((point) => point.remainingOpen)), 4)
  const barScale = niceScale(Math.max(1, ...visible.flatMap((point) => [point.opened, point.closed])), 4)
  const toYLine = (value: number) => pad.t + chartH - (value / lineScale.max) * chartH
  const toYBar = (value: number) => pad.t + chartH - (value / barScale.max) * chartH
  const points = visible.map((point, index) => ({
    x: pad.l + index * group + group / 2,
    y: toYLine(point.remainingOpen),
    point,
  }))
  const hovered = hoveredIndex === null ? null : visible[hoveredIndex]
  const hoveredPoint = hoveredIndex === null ? null : points[hoveredIndex]

  return (
    <div className="chart-shell">
      <div className="chart-topline">
        <div className="chart-legend">
          <span className="legend-remaining">Remaining Open</span>
          <span className="legend-opened">Opened</span>
          <span className="legend-closed">Closed</span>
        </div>
        <span className="chart-window">Last {visible.length} reporting weeks</span>
      </div>
      <svg viewBox={`0 0 ${width} ${height}`} className="chart-svg" role="img" aria-label="Issues by Work Week">
        <defs>
          <linearGradient id="remaining-area" x1="0" x2="0" y1="0" y2="1">
            <stop offset="0%" stopColor="#c2870b" stopOpacity="0.16" />
            <stop offset="100%" stopColor="#c2870b" stopOpacity="0" />
          </linearGradient>
        </defs>
        {barScale.ticks.map((tick) => (
          <g key={`bar-axis-${tick}`}>
            <line x1={pad.l} x2={width - pad.r} y1={toYBar(tick)} y2={toYBar(tick)} className="grid-line" />
            <text className="tick-label" x={pad.l - 9} y={toYBar(tick) + 4} textAnchor="end">{chartValue(tick)}</text>
          </g>
        ))}
        {lineScale.ticks.map((tick) => (
          <text key={`line-axis-${tick}`} className="tick-label remaining-axis-label" x={width - pad.r + 9} y={toYLine(tick) + 4}>{chartValue(tick)}</text>
        ))}
        {visible.map((point, index) => {
          const isLatest = index === visible.length - 1
          return (
            <g key={`hit-${point.workWeek}`} onMouseEnter={() => setHoveredIndex(index)} onMouseLeave={() => setHoveredIndex(null)}>
              {isLatest && <rect className="report-week-band" x={pad.l + index * group} y={pad.t} width={group} height={chartH} rx="6" />}
              <rect className="chart-hit-area" x={pad.l + index * group} y={pad.t} width={group} height={chartH} />
            </g>
          )
        })}
        {visible.map((point, index) => {
          const baseX = pad.l + index * group + group * 0.17
          const openedH = (point.opened / barScale.max) * chartH
          const closedH = (point.closed / barScale.max) * chartH
          const barWidth = Math.max(3, group * 0.25)
          const openedInside = openedH >= 9
          const closedInside = closedH >= 9
          return (
            <g key={point.workWeek}>
              <rect className="bar opened" x={baseX} y={pad.t + chartH - openedH} width={barWidth} height={openedH} rx="2.5" />
              <rect className="bar closed" x={baseX + group * 0.34} y={pad.t + chartH - closedH} width={barWidth} height={closedH} rx="2.5" />
              {point.opened > 0 && <text className={cx('bar-value', 'dense', openedInside && 'inside')} x={baseX + barWidth / 2} y={openedInside ? pad.t + chartH - 2.5 : toYBar(point.opened) - 2.5} textAnchor="middle">{point.opened}</text>}
              {point.closed > 0 && <text className={cx('bar-value', 'dense', closedInside && 'inside')} x={baseX + group * 0.34 + barWidth / 2} y={closedInside ? pad.t + chartH - 2.5 : toYBar(point.closed) - 2.5} textAnchor="middle">{point.closed}</text>}
              <text
                className="week-axis-label"
                transform={`translate(${pad.l + index * group + group / 2} ${height - 10}) rotate(-40)`}
                textAnchor="end"
              >
                {workWeekNumber(point.workWeek)}
              </text>
            </g>
          )
        })}
        <path d={areaPath(points, pad.t + chartH)} className="remaining-area" />
        <path d={smoothPath(points)} className="remaining-line" fill="none" />
        {points.map(({ x, y, point }, index) => (
          <g key={point.workWeek}>
            <circle cx={x} cy={y} r={index === hoveredIndex ? 4.5 : 3} className="remaining-dot" />
            {(index % 3 === 0 || index === points.length - 1) && (
              <text className="line-value" x={x} y={Math.max(pad.t + 7, y - 4.5)} textAnchor="middle">{point.remainingOpen}</text>
            )}
          </g>
        ))}
      </svg>
      {hovered && hoveredPoint && (
        <ChartTooltip
          title={hovered.workWeek}
          entries={[`Remaining ${hovered.remainingOpen}`, `Opened ${hovered.opened}`, `Closed ${hovered.closed}`]}
          xPercent={(hoveredPoint.x / width) * 100}
        />
      )}
    </div>
  )
}

function OpenAgingChart({ data }: { data: AgingBucket[] }) {
  const [hoveredIndex, setHoveredIndex] = useState<number | null>(null)
  const { max, ticks } = niceScale(Math.max(1, ...data.map((bucket) => bucket.count)), 4)
  const width = 700
  const height = 150
  const pad = { l: 48, r: 18, t: 22, b: 38 }
  const chartW = width - pad.l - pad.r
  const chartH = height - pad.t - pad.b
  const group = chartW / Math.max(1, data.length)
  const barWidth = Math.min(72, group * 0.56)
  const toY = (value: number) => pad.t + chartH - (value / max) * chartH
  const hovered = hoveredIndex === null ? null : data[hoveredIndex]
  return (
    <div className="chart-shell mini-chart-shell open-aging-shell">
      <svg className="mini-chart" viewBox={`0 0 ${width} ${height}`} role="img" aria-label="Open Issue Aging by Weeks">
        {ticks.map((tick) => (
          <g key={tick}>
            <line x1={pad.l} x2={width - pad.r} y1={toY(tick)} y2={toY(tick)} className="grid-line" />
            <text className="tick-label" x={pad.l - 8} y={toY(tick) + 4} textAnchor="end">{chartValue(tick)}</text>
          </g>
        ))}
        {data.map((bucket, index) => {
          const x = pad.l + index * group + (group - barWidth) / 2
          const barHeight = (bucket.count / max) * chartH
          const y = pad.t + chartH - barHeight
          const labelInside = barHeight >= 24
          return (
            <g key={bucket.label} onMouseEnter={() => setHoveredIndex(index)} onMouseLeave={() => setHoveredIndex(null)}>
              <rect className="chart-hit-area" x={pad.l + index * group} y={pad.t} width={group} height={chartH} />
              <rect
                className={cx('bar', 'open-aging', index === hoveredIndex && 'active')}
                style={{ fill: bucket.color }}
                x={x}
                y={y}
                width={barWidth}
                height={Math.max(bucket.count > 0 ? 2 : 0, barHeight)}
                rx="5"
              />
              <text
                className={cx('open-aging-value', labelInside && 'inside')}
                x={x + barWidth / 2}
                y={labelInside ? y + 17 : Math.max(14, y - 7)}
                textAnchor="middle"
              >
                {bucket.count}
              </text>
              <text className="open-aging-bucket" x={x + barWidth / 2} y={height - 20} textAnchor="middle">{bucket.label}</text>
            </g>
          )
        })}
        <text className="axis-title open-aging-y-title" x="13" y={pad.t + chartH / 2} textAnchor="middle" transform={`rotate(-90 13 ${pad.t + chartH / 2})`}>Open issues</text>
        <text className="axis-title open-aging-x-title" x={pad.l + chartW / 2} y={height - 4} textAnchor="middle">Weeks</text>
      </svg>
      {hovered && (
        <ChartTooltip
          compact
          title={`${hovered.label} weeks`}
          entries={[`Open issues ${hovered.count}`]}
          xPercent={((hoveredIndex ?? 0) + 0.5) / data.length * 100}
        />
      )}
    </div>
  )
}

function AgingChart({ data }: { data: AgingBucket[] }) {
  const total = data.reduce((sum, bucket) => sum + bucket.count, 0)
  return (
    <div className="aging-chart">
      <div className="aging-stack" aria-label="Issue aging distribution">
        {data.map((bucket) => (
          <i
            key={bucket.label}
            style={{ width: `${total ? (bucket.count / total) * 100 : 0}%`, background: bucket.color }}
            title={`${bucket.label}: ${bucket.count}`}
          />
        ))}
      </div>
      {data.map((bucket) => (
        <div className="aging-row" key={bucket.label}>
          <span><i style={{ background: bucket.color }} />{bucket.label}</span>
          <div>
            <i style={{ width: `${total ? (bucket.count / total) * 100 : 0}%`, background: bucket.color }} />
          </div>
          <strong>{bucket.count}<em>{total ? Math.round((bucket.count / total) * 100) : 0}%</em></strong>
        </div>
      ))}
    </div>
  )
}

function ElectricalChart({ data }: { data: ElectricalPoint[] }) {
  const [hoveredIndex, setHoveredIndex] = useState<number | null>(null)
  const visible = data.slice(-30)
  const { max, ticks } = niceScale(Math.max(1, ...visible.map((point) => point.finals)))
  const width = 620
  const height = 360
  const pad = { l: 44, r: 22, t: 28, b: 42 }
  const chartW = width - pad.l - pad.r
  const chartH = height - pad.t - pad.b
  const step = chartW / Math.max(1, visible.length - 1)
  const toY = (value: number) => pad.t + chartH - (value / max) * chartH
  const points = visible.map((point, index) => ({ x: pad.l + index * step, y: toY(point.finals) }))
  const hovered = hoveredIndex === null ? null : visible[hoveredIndex]
  return (
    <div className="chart-shell field-chart-shell">
      <svg className="chart-svg" viewBox={`0 0 ${width} ${height}`} role="img" aria-label="Electrical Inspections by Work Week">
        <defs>
          <linearGradient id="electrical-area" x1="0" x2="0" y1="0" y2="1">
            <stop offset="0%" stopColor="#2e5aac" stopOpacity="0.16" />
            <stop offset="100%" stopColor="#2e5aac" stopOpacity="0" />
          </linearGradient>
        </defs>
        {ticks.map((tick) => (
          <g key={tick}>
            <line x1={pad.l} x2={width - pad.r} y1={toY(tick)} y2={toY(tick)} className="grid-line" />
            <text className="tick-label" x={pad.l - 9} y={toY(tick) + 4} textAnchor="end">{chartValue(tick)}</text>
          </g>
        ))}
        <path d={areaPath(points, pad.t + chartH)} className="electrical-area" />
        <path d={smoothPath(points)} className="electrical-line" fill="none" />
        {visible.map((point, index) => {
          const { x, y } = points[index]
          const isActive = index === hoveredIndex
          const issueY = Math.min(pad.t + chartH - 11, y + 21)
          return (
            <g key={point.workWeek} onMouseEnter={() => setHoveredIndex(index)} onMouseLeave={() => setHoveredIndex(null)}>
              <rect className="chart-hit-area" x={x - step / 2} y={pad.t} width={step} height={chartH} />
              <circle cx={x} cy={y} r={isActive ? 5 : 3.5} className="electrical-dot" />
              <text className="electrical-point-label" x={x} y={Math.max(12, y - 10)} textAnchor="middle">{point.finals}</text>
              {point.issuesFound > 0 && (
                <g className="issue-badge" transform={`translate(${x} ${issueY})`}>
                  <circle r="8" />
                  <text textAnchor="middle" y="3">{point.issuesFound}</text>
                </g>
              )}
              {(index % 5 === 0 || index === visible.length - 1) && (
                <text className="week-axis-label" x={x} y={height - 22} textAnchor="middle">
                  {workWeekNumber(point.workWeek)}
                </text>
              )}
            </g>
          )
        })}
      </svg>
      {hovered && (
        <ChartTooltip
          title={hovered.workWeek}
          entries={[`Final inspections ${hovered.finals}`, `Issues found ${hovered.issuesFound}`]}
          xPercent={visible.length > 1 ? ((hoveredIndex ?? 0) / (visible.length - 1)) * 100 : 50}
        />
      )}
    </div>
  )
}

function WeldingChart({ data, reportWeek }: { data: WeldingPoint[]; reportWeek: string }) {
  const [hoveredIndex, setHoveredIndex] = useState<number | null>(null)
  const visible = data.slice(-30)
  const { max, ticks } = niceScale(Math.max(1, ...visible.map((point) => point.total)))
  const width = 620
  const height = 360
  const pad = { l: 44, r: 48, t: 34, b: 42 }
  const chartW = width - pad.l - pad.r
  const chartH = height - pad.t - pad.b
  const group = chartW / Math.max(visible.length, 1)
  const barW = Math.min(24, group * 0.56)
  const toY = (value: number) => pad.t + chartH - (value / max) * chartH
  const toRateY = (value: number) => pad.t + chartH - (value / 100) * chartH
  const rateTicks = [0, 25, 50, 75, 100]
  const ratePoints = visible
    .map((point, index) => ({
      point,
      x: pad.l + index * group + group / 2,
      y: toRateY(point.signoffRate),
    }))
    .filter(({ point }) => point.total > 0)
  const reportRatePoint = ratePoints.find(({ point }) => point.workWeek === reportWeek)
  const reportRateLabelY = (() => {
    if (!reportRatePoint) return null
    const above = Math.max(pad.t + 10, reportRatePoint.y - 10)
    const below = Math.min(pad.t + chartH - 5, reportRatePoint.y + 17)
    const totalY = toY(reportRatePoint.point.total)
    const issueY = Math.max(pad.t + 10, totalY - 20)
    const reservedY = reportRatePoint.point.issuesCreated > 0 ? issueY : Math.max(pad.t + 9, totalY - 5)
    return Math.abs(above - reservedY) < 15 ? below : above
  })()
  const hovered = hoveredIndex === null ? null : visible[hoveredIndex]
  return (
    <div className="chart-shell field-chart-shell">
      <svg className="chart-svg" viewBox={`0 0 ${width} ${height}`} role="img" aria-label="Welding Signoffs by Work Week">
        {ticks.map((tick) => (
          <g key={tick}>
            <line x1={pad.l} x2={width - pad.r} y1={toY(tick)} y2={toY(tick)} className="grid-line" />
            <text className="tick-label" x={pad.l - 9} y={toY(tick) + 4} textAnchor="end">{chartValue(tick)}</text>
          </g>
        ))}
        {rateTicks.map((tick) => (
          <text className="tick-label signoff-axis-label" key={tick} x={width - pad.r + 9} y={toRateY(tick) + 4}>{tick}%</text>
        ))}
        {visible.map((point, index) => {
          const x = pad.l + index * group + (group - barW) / 2
          const totalH = (point.total / max) * chartH
          const signedH = (point.signed / max) * chartH
          const totalY = pad.t + chartH - totalH
          const signedY = pad.t + chartH - signedH
          return (
            <g key={point.workWeek} onMouseEnter={() => setHoveredIndex(index)} onMouseLeave={() => setHoveredIndex(null)}>
              <rect className="chart-hit-area" x={pad.l + index * group} y={pad.t} width={group} height={chartH} />
              <rect className="bar track" x={x} y={totalY} width={barW} height={totalH} rx="3" />
              <rect className="bar closed" x={x} y={signedY} width={barW} height={signedH} rx="3" />
              {point.total > 0 && (
                <text className="weld-total-label" x={x + barW / 2} y={Math.max(pad.t + 9, totalY - 5)} textAnchor="middle">{point.total}</text>
              )}
              {point.signed > 0 && (
                <text
                  className={cx('weld-signed-label', signedH >= 16 && 'inside')}
                  x={x + barW / 2}
                  y={signedH >= 16 ? signedY + 12 : signedY - 4}
                  textAnchor="middle"
                >
                  {point.signed}
                </text>
              )}
              {(index % 5 === 0 || index === visible.length - 1) && (
                <text className="week-axis-label" x={pad.l + index * group + group / 2} y={height - 22} textAnchor="middle">
                  {workWeekNumber(point.workWeek)}
                </text>
              )}
            </g>
          )
        })}
        <line className="signoff-baseline" x1={pad.l} x2={width - pad.r} y1={toRateY(10)} y2={toRateY(10)} />
        <path className="signoff-rate-line" d={smoothPath(ratePoints)} fill="none" />
        {ratePoints.map(({ point, x, y }) => (
          <g key={`rate-${point.workWeek}`}>
            <circle className="signoff-rate-dot" cx={x} cy={y} r="3.5" />
            {point.workWeek === reportWeek && reportRateLabelY !== null && (
              <>
                <rect className="signoff-rate-label-bg" x={x - 14} y={reportRateLabelY - 9} width="28" height="12" rx="4" />
                <text className="signoff-rate-label" x={x} y={reportRateLabelY} textAnchor="middle">{percent(point.signoffRate, 0)}</text>
              </>
            )}
          </g>
        ))}
        {visible.map((point, index) => {
          if (point.issuesCreated <= 0) return null
          const x = pad.l + index * group + group / 2
          const totalY = toY(point.total)
          const issueY = Math.max(pad.t + 10, totalY - 20)
          return (
            <g className="issue-badge" key={`issue-${point.workWeek}`} transform={`translate(${x} ${issueY})`}>
              <circle r="8" />
              <text textAnchor="middle" y="3">{point.issuesCreated}</text>
            </g>
          )
        })}
        <text className="signoff-baseline-label" x={width - 4} y={toRateY(10) - 5} textAnchor="end">10% baseline</text>
        <text className="axis-title" x={12} y={pad.t + chartH / 2} textAnchor="middle" transform={`rotate(-90 12 ${pad.t + chartH / 2})`}>Weld count</text>
        <text className="axis-title signoff-axis-label" x={width - 10} y={pad.t + chartH / 2} textAnchor="middle" transform={`rotate(90 ${width - 10} ${pad.t + chartH / 2})`}>Sign-off rate</text>
      </svg>
      {hovered && (
        <ChartTooltip
          title={hovered.workWeek}
          entries={[`Signed ${hovered.signed} of ${hovered.total}`, `Sign-off ${percent(hovered.signoffRate, 1)}`, `Issues created ${hovered.issuesCreated}`]}
          xPercent={visible.length > 1 ? ((hoveredIndex ?? 0) / (visible.length - 1)) * 100 : 50}
        />
      )}
    </div>
  )
}

function SlideShell({
  children,
  title,
  icon,
  exportable,
  meta,
  hideHeader,
  className,
  headerExtra,
}: {
  children: React.ReactNode
  title: string
  icon: React.ReactNode
  exportable?: boolean
  meta?: string
  hideHeader?: boolean
  className?: string
  headerExtra?: React.ReactNode
}) {
  return (
    <section
      className={cx('slide-frame', hideHeader && 'headerless', className)}
      aria-label={title}
      {...(exportable ? { 'data-export-slide': true } : {})}
    >
      <div className="gc-guard top" />
      <div className="gc-guard bottom" />
      <div className="slide-safe">
        {!hideHeader && (
          <header className="slide-header">
            <div>
              <span>{icon}</span>
              <h2>{title}</h2>
            </div>
            {headerExtra ?? (meta && <em className="slide-meta">{meta}</em>)}
          </header>
        )}
        {children}
      </div>
    </section>
  )
}

function IssueDetailHeader({
  pageIndex,
  pageCount,
  onPreviousPage,
  onNextPage,
}: {
  pageIndex: number
  pageCount: number
  onPreviousPage?: () => void
  onNextPage?: () => void
}) {
  const interactive = Boolean(onPreviousPage && onNextPage)
  return (
    <div className="issue-detail-header-tools">
      {pageCount > 1 && (
        <div className="issue-page-controls" aria-label={`Issue detail page ${pageIndex + 1} of ${pageCount}`}>
          {interactive && (
            <button
              aria-label="Previous issue page"
              disabled={pageIndex === 0}
              onClick={onPreviousPage}
              title="Previous issue page"
              type="button"
            >
              <ChevronLeft size={14} />
            </button>
          )}
          <strong>{pageIndex + 1} / {pageCount}</strong>
          {interactive && (
            <button
              aria-label="Next issue page"
              disabled={pageIndex >= pageCount - 1}
              onClick={onNextPage}
              title="Next issue page"
              type="button"
            >
              <ChevronRight size={14} />
            </button>
          )}
        </div>
      )}
      <div className="issue-row-legend" aria-label="Issue row legend">
        <span><i className="reporting-week" />Reporting Week</span>
        <span><i className="recent-open" />Most Recent Open Issues</span>
      </div>
    </div>
  )
}

function OverviewSlide({ report, exportable }: { report: ReturnType<typeof buildReportModel>; exportable?: boolean }) {
  return (
    <SlideShell title="Weekly QA/QC Report" icon={<Gauge size={18} />} exportable={exportable} hideHeader>
      <KpiZone report={report} />
      <div className="overview-grid">
        <article className="panel chart-panel overview-trend-panel">
          <div className="panel-title">
            <LineChart size={16} />
            <h3>Issues by Work Week</h3>
          </div>
          <RangeChart data={report.issueTrend} />
        </article>
        <article className="panel chart-panel">
          <div className="panel-title">
            <History size={16} />
            <h3>Open Issue Aging</h3>
          </div>
          <OpenAgingChart data={report.openAging} />
        </article>
        <article className="panel chart-panel">
          <div className="panel-title">
            <CalendarClock size={16} />
            <h3>Overall Issue Aging (Opened &amp; Closed)</h3>
          </div>
          <AgingChart data={report.aging} />
        </article>
      </div>
    </SlideShell>
  )
}

function IssueTableSlide({
  report,
  exportable,
  rows,
  summaryRows,
  pageIndex = 0,
  pageCount = 1,
  onPreviousPage,
  onNextPage,
}: {
  report: ReturnType<typeof buildReportModel>
  exportable?: boolean
  rows?: IssueDetailRow[]
  summaryRows?: IssueDetailRow[]
  pageIndex?: number
  pageCount?: number
  onPreviousPage?: () => void
  onNextPage?: () => void
}) {
  const allRows = summaryRows ?? report.issueTable
  const groupCounts = allRows.reduce<Record<string, number>>((acc, row) => {
    acc[row.group] = (acc[row.group] ?? 0) + 1
    return acc
  }, {})
  const visibleRows = rows ?? report.issueTable
  const detailCards: Array<{ group: IssueDetailRow['group']; label: string }> = [
    report.activeFilters.oac
      ? {
          group: 'Opened in Report Week',
          label: report.reportingMode === 'justine' ? 'Issues Opened During Reporting Period' : 'Issues Opened During Report Week',
        }
      : { group: 'Open Carryover', label: 'Issues Remaining Open' },
    {
      group: 'Closed in Report Week',
      label: report.reportingMode === 'justine' ? 'Issues Closed During Reporting Period' : 'Issues Closed During Report Week',
    },
    {
      group: 'Opened + Closed in Report Week',
      label: report.reportingMode === 'justine' ? 'Opened + Closed Within Reporting Period' : 'Opened + Closed Within Report Week',
    },
    { group: 'Closed This Week', label: 'Issues Closed During Current Week' },
  ]
  return (
    <SlideShell
      title="BIM Issues Detail"
      icon={<Table2 size={18} />}
      exportable={exportable}
      className="issue-detail-slide"
      headerExtra={(
        <IssueDetailHeader
          pageIndex={pageIndex}
          pageCount={pageCount}
          onPreviousPage={onPreviousPage}
          onNextPage={onNextPage}
        />
      )}
    >
      <div className="detail-metrics">
        {detailCards.map(({ group, label }) => (
          <div className="metric-chip" key={group}>
            <strong>{compactNumber(groupCounts[group] ?? 0)}</strong>
            <span>{label}</span>
          </div>
        ))}
      </div>
      <div className={cx('issue-table-wrap', exportable && 'export-page')}>
        <table className="issue-table">
          <thead>
            <tr>
              {['ID', 'Subtype', 'Status', 'Title', 'Contractor', 'Work Week', 'Created On', 'Work Week Closed', 'Due Date'].map((header) => (
                <th key={header}>{header}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {visibleRows.map((row, index) => (
              <tr
                className={cx(row.group === 'Open for Discussion' && 'open-discussion-row')}
                key={`${row.id}-${row.workWeekClosed}-${index}`}
                title={row.group === 'Open for Discussion' ? 'Open for discussion' : undefined}
              >
                <td>{row.id}</td>
                <td>{row.subtype}</td>
                <td>
                  <span className={cx('status-pill', row.status.toLowerCase().replace(/\s+/g, '-'))}>{row.status}</span>
                </td>
                <td>{row.title}</td>
                <td>{row.contractor}</td>
                <td>{row.workWeek}</td>
                <td>{row.createdOn}</td>
                <td>{row.workWeekClosed}</td>
                <td>{row.dueDate}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </SlideShell>
  )
}

function FieldSlide({ report, exportable }: { report: ReturnType<typeof buildReportModel>; exportable?: boolean }) {
  const chips = [
    ['Electrical Finals', compactNumber(report.summary.electricalFinals), deltaLabel(report.summary.deltas.electricalFinals), ClipboardCheck],
    ['Electrical Issues Found', compactNumber(report.summary.electricalIssuesFound), deltaLabel(report.summary.deltas.electricalIssuesFound), AlertCircle],
    ['Total Welds', compactNumber(report.summary.weldsChecked), deltaLabel(report.summary.deltas.weldsChecked), Flame],
    ['Welds Signed', compactNumber(report.summary.weldsSigned), deltaLabel(report.summary.deltas.weldsSigned), CheckCircle2],
  ] as const
  return (
    <SlideShell
      title="Inspections & Welding Signoffs"
      icon={<ClipboardCheck size={18} />}
      exportable={exportable}
      meta={`Current ${report.currentWeek.label}`}
    >
      <div className="field-chip-row">
        {chips.map(([label, value, delta, Icon]) => (
          <article className="field-chip" key={label}>
            <Icon size={15} />
            <span>{label}</span>
            <strong>{value}</strong>
            <em>{delta}</em>
          </article>
        ))}
        <article className="field-chip signoff-summary-chip">
          <div className="field-chip-heading">
            <Gauge size={15} />
            <span>Sign-off %</span>
          </div>
          <div className="signoff-summary-grid">
            <div>
              <small>Overall avg</small>
              <strong>{percent(report.summary.overallSignoffRate, 1)}</strong>
            </div>
            <div>
              <small>{report.periodLabel}</small>
              <strong>{percent(report.summary.reportWeekSignoffRate, 1)}</strong>
              <em>{deltaLabel(report.summary.deltas.reportWeekSignoffRate)}</em>
            </div>
          </div>
        </article>
      </div>
      <div className="field-chart-grid">
        <article className="panel chart-panel">
          <div className="panel-title">
            <Zap size={16} />
            <h3>Electrical Inspections by Work Week</h3>
          </div>
          <div className="chart-legend">
            <span className="legend-final">Final Inspections</span>
            <span className="legend-issue">Issue Found</span>
          </div>
          <ElectricalChart data={report.electrical} />
        </article>
        <article className="panel chart-panel">
          <div className="panel-title">
            <Flame size={16} />
            <h3>Welding Signoffs by Work Week</h3>
          </div>
          <div className="chart-legend">
            <span className="legend-closed">Signed</span>
            <span className="legend-track">Total welds</span>
            <span className="legend-signoff">Sign-off %</span>
            <span className="legend-issue">Issue created</span>
            <span className="legend-baseline">10% baseline</span>
          </div>
          <WeldingChart data={report.welding} reportWeek={report.reportWeek.label} />
        </article>
      </div>
    </SlideShell>
  )
}

function ConnectFirst({
  imports,
  onChoose,
  onChooseFolder,
  onDrop,
  onPreview,
  importing,
}: {
  imports: Partial<Record<SheetRole, ImportedSheetFile>>
  onChoose: () => void
  onChooseFolder: () => void
  onDrop: (dataTransfer: DataTransfer) => void
  onPreview: () => void
  importing: boolean
}) {
  const [dragging, setDragging] = useState(false)
  const importCount = Object.keys(imports).length

  return (
    <section
      className={cx('connect-first', dragging && 'dragging')}
      onDragEnter={(event) => {
        event.preventDefault()
        setDragging(true)
      }}
      onDragOver={(event) => event.preventDefault()}
      onDragLeave={(event) => {
        if (event.currentTarget === event.target) setDragging(false)
      }}
      onDrop={(event) => {
        event.preventDefault()
        setDragging(false)
        onDrop(event.dataTransfer)
      }}
    >
      <div className="connect-first-mark">
        <Archive size={34} />
      </div>
      <h2>Drop a ZIP, folder, or weekly exports</h2>
      <p>{importCount}/4 required files ready</p>
      <div className="connect-first-actions">
        <button className="button primary" type="button" onClick={onChoose} disabled={importing}>
          {importing ? <RefreshCw size={16} className="spin" /> : <Archive size={16} />}
          Choose ZIP or Files
        </button>
        <button className="button secondary" type="button" onClick={onChooseFolder} disabled={importing}>
          <FolderOpen size={16} />
          Choose Folder
        </button>
        <button className="button secondary" type="button" onClick={onPreview} disabled={importing}>
          <LineChart size={16} />
          Preview Layout
        </button>
      </div>
      <div className="connect-first-grid">
        {(Object.keys(ROLE_CONFIG) as SheetRole[]).map((role) => {
          const file = imports[role]
          return (
            <span className={cx(file && 'ready')} key={role}>
              {file ? <CheckCircle2 size={14} /> : <FileSpreadsheet size={14} />}
              {ROLE_CONFIG[role].label}
            </span>
          )
        })}
      </div>
      <small>.zip, .xls, .xlsx, .csv, or a full folder</small>
    </section>
  )
}

function ImportPanel({
  imports,
  onChoose,
  onChooseFolder,
  onRemove,
  onClear,
  importing,
}: {
  imports: Partial<Record<SheetRole, ImportedSheetFile>>
  onChoose: () => void
  onChooseFolder: () => void
  onRemove: (role: SheetRole) => void
  onClear: () => void
  importing: boolean
}) {
  const importCount = Object.keys(imports).length
  return (
    <aside className="sheet-panel">
      <div className="import-panel-header">
        <div className="panel-title">
          <FileSpreadsheet size={16} />
          <h3>Report Files</h3>
        </div>
        <strong>{importCount}/4</strong>
      </div>
      <button className="button primary import-button" type="button" onClick={onChoose} disabled={importing}>
        {importing ? <RefreshCw size={16} className="spin" /> : <Archive size={16} />}
        Import ZIP or Files
      </button>
      <button className="button secondary import-button folder-button" type="button" onClick={onChooseFolder} disabled={importing}>
        <FolderOpen size={16} />
        Import Folder
      </button>
      <div className="import-role-list">
        {(Object.keys(ROLE_CONFIG) as SheetRole[]).map((role) => {
          const config = ROLE_CONFIG[role]
          const Icon = config.icon
          const file = imports[role]
          return (
            <div className={cx('import-role-row', file && 'ready')} key={role}>
              <span className="import-role-icon" style={{ color: config.color }}>
                <Icon size={15} />
              </span>
              <div>
                <strong>{config.label}</strong>
                <span title={file?.fileName}>{file ? file.fileName : 'Not imported'}</span>
                {file && <small>{file.rowCount.toLocaleString()} rows</small>}
              </div>
              {file ? (
                <button className="icon-button compact" type="button" onClick={() => onRemove(role)} aria-label={`Remove ${config.label}`} title="Remove file">
                  <Trash2 size={14} />
                </button>
              ) : (
                <FileSpreadsheet className="missing-file-icon" size={15} />
              )}
            </div>
          )
        })}
      </div>
      {importCount > 0 && (
        <button className="clear-imports" type="button" onClick={onClear}>
          <Trash2 size={14} />
          Clear all files
        </button>
      )}
      <div className="local-data-note">
        <ShieldCheck size={15} />
        <span>Spreadsheet data remains in this browser session.</span>
      </div>
    </aside>
  )
}

function ManualIssueFileSlot({
  kind,
  title,
  workbook,
  loading,
  onChoose,
  onFile,
  onClear,
}: {
  kind: ManualIssueWorkbookKind
  title: string
  workbook: PreparedIssueWorkbook | null
  loading: boolean
  onChoose: () => void
  onFile: (file: File) => void
  onClear: () => void
}) {
  const [dragging, setDragging] = useState(false)
  return (
    <div
      className={cx('manual-file-slot', workbook && 'ready', dragging && 'dragging')}
      onDragEnter={(event) => {
        event.preventDefault()
        setDragging(true)
      }}
      onDragOver={(event) => event.preventDefault()}
      onDragLeave={(event) => {
        if (event.currentTarget === event.target) setDragging(false)
      }}
      onDrop={(event) => {
        event.preventDefault()
        setDragging(false)
        const file = event.dataTransfer.files[0]
        if (file) onFile(file)
      }}
    >
      <div className="manual-file-icon">
        {workbook ? <CheckCircle2 size={20} /> : <FileSpreadsheet size={20} />}
      </div>
      <div className="manual-file-copy">
        <span>{kind === 'current' ? 'Source' : 'New export'}</span>
        <strong>{title}</strong>
        {workbook ? (
          <small title={workbook.fileName}>
            {workbook.fileName} · {workbook.rowCount.toLocaleString()} rows
          </small>
        ) : (
          <small>.xlsx, .xls, or .csv</small>
        )}
      </div>
      {workbook ? (
        <button className="icon-button compact" type="button" onClick={onClear} aria-label={`Remove ${title}`} title="Remove file">
          <X size={15} />
        </button>
      ) : (
        <button className="button secondary" type="button" onClick={onChoose} disabled={loading}>
          {loading ? <RefreshCw className="spin" size={16} /> : <Archive size={16} />}
          Choose file
        </button>
      )}
    </div>
  )
}

const ENRICHED_FIELD_LABELS: Record<EnrichedIssueField, string> = {
  contractor: 'Contractor',
  discipline: 'Discipline',
  createdBy: 'BIM owner',
  createdOn: 'Created date',
  closedOn: 'Closed date',
}

function enrichedFieldList(fields: EnrichedIssueField[]): string {
  return fields.map((field) => ENRICHED_FIELD_LABELS[field]).join(' + ')
}

function drivingDateSummary(change: ManualIssueEnrichment): string {
  const dates = [
    change.filledFields.includes('createdOn') && change.createdOn
      ? `Created ${formatDate(change.createdOn)}`
      : '',
    change.filledFields.includes('closedOn') && change.closedOn
      ? `Closed ${formatDate(change.closedOn)}`
      : '',
  ].filter(Boolean)
  return dates.join(' · ') || 'No date change'
}

function ManualIssuesUpdate() {
  const currentInputRef = useRef<HTMLInputElement>(null)
  const accInputRef = useRef<HTMLInputElement>(null)
  const [current, setCurrent] = useState<PreparedIssueWorkbook | null>(null)
  const [acc, setAcc] = useState<PreparedIssueWorkbook | null>(null)
  const [loading, setLoading] = useState<ManualIssueWorkbookKind | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [downloadedFile, setDownloadedFile] = useState<string | null>(null)
  const [summaryOpen, setSummaryOpen] = useState(false)

  const analysis = useMemo(
    () => (current && acc ? reconcileIssueRows(current, acc) : null),
    [current, acc],
  )
  const issueChanges = useMemo(
    () => analysis?.changes ?? [],
    [analysis],
  )

  useEffect(() => {
    if (analysis) setSummaryOpen(true)
  }, [analysis])

  async function loadFile(file: File, kind: ManualIssueWorkbookKind): Promise<void> {
    setLoading(kind)
    setError(null)
    setDownloadedFile(null)
    try {
      const workbook = await prepareIssueWorkbook(file, kind)
      if (kind === 'current') setCurrent(workbook)
      else setAcc(workbook)
    } catch (err) {
      setError(err instanceof Error ? err.message : 'The issue workbook could not be loaded.')
    } finally {
      setLoading(null)
    }
  }

  function downloadUpdatedLog(): void {
    if (!acc || !analysis) return
    setError(null)
    try {
      const output = buildUpdatedIssueWorkbook(acc, analysis)
      downloadUpdatedIssueWorkbook(output)
      setDownloadedFile(output.fileName)
    } catch (err) {
      setError(err instanceof Error ? err.message : 'The enriched ACC export could not be created.')
    }
  }

  return (
    <section className="manual-update-workspace">
      <input
        ref={currentInputRef}
        className="file-input"
        type="file"
        accept=".xls,.xlsx,.csv"
        onChange={(event) => {
          const file = event.target.files?.[0]
          if (file) void loadFile(file, 'current')
          event.target.value = ''
        }}
      />
      <input
        ref={accInputRef}
        className="file-input"
        type="file"
        accept=".xls,.xlsx,.csv"
        onChange={(event) => {
          const file = event.target.files?.[0]
          if (file) void loadFile(file, 'acc')
          event.target.value = ''
        }}
      />

      <div className="manual-update-heading">
        <div>
          <span>ACC METADATA ENRICHMENT</span>
          <h2>Manual Issues Update</h2>
        </div>
        <div className="secure-mark">
          <ShieldCheck size={14} />
          Local processing
        </div>
      </div>

      {error && (
        <div className="error-banner manual-error">
          <AlertCircle size={16} />
          <span>{error}</span>
          <button type="button" onClick={() => setError(null)}>Dismiss</button>
        </div>
      )}

      <div className="manual-file-grid">
        <ManualIssueFileSlot
          kind="current"
          title="Historical BIM Issues Log"
          workbook={current}
          loading={loading === 'current'}
          onChoose={() => currentInputRef.current?.click()}
          onFile={(file) => void loadFile(file, 'current')}
          onClear={() => {
            setCurrent(null)
            setDownloadedFile(null)
          }}
        />
        <ManualIssueFileSlot
          kind="acc"
          title="ACC Export to Enrich"
          workbook={acc}
          loading={loading === 'acc'}
          onChoose={() => accInputRef.current?.click()}
          onFile={(file) => void loadFile(file, 'acc')}
          onClear={() => {
            setAcc(null)
            setDownloadedFile(null)
          }}
        />
      </div>

      {analysis ? (
        <>
          <div className="manual-summary-grid">
            <article>
              <span>ACC export</span>
              <strong>{analysis.accRows.toLocaleString()}</strong>
              <small>rows preserved</small>
            </article>
            <article>
              <span>Matched IDs</span>
              <strong>{analysis.matchedRows.toLocaleString()}</strong>
              <small>reference rows found</small>
            </article>
            <article className="new-issues">
              <span>Contractor fills</span>
              <strong>{analysis.filledContractors.toLocaleString()}</strong>
              <small>blank ACC cells</small>
            </article>
            <article>
              <span>Discipline fills</span>
              <strong>{analysis.filledDisciplines.toLocaleString()}</strong>
              <small>blank ACC cells</small>
            </article>
            <article className="new-issues">
              <span>Driving dates</span>
              <strong>{(analysis.enrichedCreatedDates + analysis.enrichedClosedDates).toLocaleString()}</strong>
              <small>
                {analysis.enrichedCreatedDates.toLocaleString()} created ·{' '}
                {analysis.enrichedClosedDates.toLocaleString()} closed
              </small>
            </article>
          </div>

          <div className="manual-results-panel">
            <div className="manual-results-header">
              <div>
                <h3>ACC Enrichment Preview</h3>
                <p>
                  {analysis.unmatchedRows.toLocaleString()} ACC IDs not found in the reference ·{' '}
                  {analysis.missingIdRows.toLocaleString()} rows without IDs left untouched ·{' '}
                  {analysis.duplicateCurrentIds.toLocaleString()} duplicate reference IDs ignored
                </p>
              </div>
              <div className="manual-results-actions">
                <button className="button secondary" type="button" onClick={() => setSummaryOpen(true)}>
                  <Table2 size={16} />
                  Review Fills
                </button>
                <button
                  className="button primary"
                  type="button"
                  onClick={downloadUpdatedLog}
                >
                  <Download size={16} />
                  Download Enriched ACC
                </button>
              </div>
            </div>

            {issueChanges.length > 0 ? (
              <div className="manual-table-scroll">
                <table className="manual-issues-table">
                  <thead>
                    <tr>
                      <th>ID</th>
                      <th>Filled fields</th>
                      <th>Contractor</th>
                      <th>Discipline</th>
                      <th>Driving dates</th>
                    </tr>
                  </thead>
                  <tbody>
                    {issueChanges.map((change) => (
                      <tr key={`${change.id}-${change.targetRow}`}>
                        <td><strong>{change.id}</strong></td>
                        <td>{enrichedFieldList(change.filledFields)}</td>
                        <td>{change.contractor || '—'}</td>
                        <td>{change.discipline || '—'}</td>
                        <td>{drivingDateSummary(change)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <div className="manual-empty-result">
                <CheckCircle2 size={24} />
                <strong>Every matching ACC row already has the latest reference metadata.</strong>
              </div>
            )}

            {downloadedFile && (
              <div className="manual-download-success">
                <CheckCircle2 size={15} />
                Downloaded {downloadedFile}
              </div>
            )}
          </div>
        </>
      ) : (
        <div className="manual-waiting-state">
          <GitBranch size={30} />
          <strong>Two files needed</strong>
          <span>The comparison appears here when both files are ready.</span>
        </div>
      )}

      <Modal open={summaryOpen && Boolean(analysis)} title="Enrichment Summary" onClose={() => setSummaryOpen(false)} wide>
        {analysis && (
          <div className="manual-summary-modal">
            <div className={cx('manual-summary-outcome', issueChanges.length === 0 && 'current')}>
              <div>
                {issueChanges.length === 0 ? <CheckCircle2 size={22} /> : <GitBranch size={22} />}
              </div>
              <span>
                <strong>
                  {issueChanges.length === 0
                    ? 'The ACC export is already complete'
                    : `${issueChanges.length.toLocaleString()} ${issueChanges.length === 1 ? 'row' : 'rows'} ready to enrich`}
                </strong>
                <small>
                  {analysis.filledContractors.toLocaleString()} Contractor cells ·{' '}
                  {analysis.filledDisciplines.toLocaleString()} Discipline cells ·{' '}
                  {(analysis.enrichedCreatedDates + analysis.enrichedClosedDates).toLocaleString()} driving dates
                </small>
              </span>
            </div>

            <div className="manual-modal-stats">
              <span><strong>{analysis.matchedRows.toLocaleString()}</strong> matching ACC rows</span>
              <span><strong>{issueChanges.length.toLocaleString()}</strong> rows receiving values</span>
              <span><strong>{analysis.unchangedMatchedRows.toLocaleString()}</strong> matched rows unchanged</span>
              <span><strong>{analysis.unmatchedRows.toLocaleString()}</strong> unmatched ACC rows preserved</span>
            </div>

            {issueChanges.length > 0 && (
              <div className="manual-modal-list">
                {issueChanges.slice(0, 8).map((change) => (
                  <div key={`summary-${change.id}-${change.targetRow}`}>
                    <strong>{change.id}</strong>
                    <span>{drivingDateSummary(change)}</span>
                    <small>{enrichedFieldList(change.filledFields)}</small>
                  </div>
                ))}
                {issueChanges.length > 8 && (
                  <p>+{(issueChanges.length - 8).toLocaleString()} more changes in the full review table</p>
                )}
              </div>
            )}

            <div className="manual-modal-actions">
              {downloadedFile && (
                <span><CheckCircle2 size={14} /> Downloaded {downloadedFile}</span>
              )}
              <button className="button secondary" type="button" onClick={() => setSummaryOpen(false)}>Close</button>
              <button
                className="button primary"
                type="button"
                onClick={downloadUpdatedLog}
              >
                <Download size={16} />
                Download Enriched ACC
              </button>
            </div>
          </div>
        )}
      </Modal>
    </section>
  )
}

export default function App() {
  const [bundle, setBundle] = useState<SheetBundle>(EMPTY_BUNDLE)
  const [filters, setFilters] = useState<ReportFilters>(() => mergeFilters(loadFilters()))
  const [workspaceTab, setWorkspaceTab] = useState<'report' | 'manual'>('report')
  const [activeSlide, setActiveSlide] = useState<'overview' | 'issues' | 'field'>('overview')
  const [issuePageIndex, setIssuePageIndex] = useState(0)
  const [imports, setImports] = useState<Partial<Record<SheetRole, ImportedSheetFile>>>({})
  const [importing, setImporting] = useState(false)
  const [exporting, setExporting] = useState(false)
  const [pdfExporting, setPdfExporting] = useState(false)
  const [sheetPanelOpen, setSheetPanelOpen] = useState(true)
  const [error, setError] = useState<string | null>(null)
  const [updateInfo, setUpdateInfo] = useState<UpdateInfo | null>(null)
  const [updateOpen, setUpdateOpen] = useState(false)
  const [updateDefaultTab, setUpdateDefaultTab] = useState<'update' | 'changelog'>('changelog')
  const [updateChecking, setUpdateChecking] = useState(false)
  const [lastUpdateCheck, setLastUpdateCheck] = useState<Date | null>(null)
  const [dataQualityReview, setDataQualityReview] = useState<DataQualityReview | null>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)
  const folderInputRef = useRef<HTMLInputElement>(null)
  const oacContractorsRef = useRef<string[]>(filters.reportingMode === 'oac' ? filters.contractors : [])
  const dismissedUpdateVersionRef = useRef<string | null>(null)
  const importCount = Object.keys(imports).length
  const filesReady = importCount === 4
  const hasReport = bundle.source === 'files' || bundle.source === 'demo'
  const showSheetPanel = !hasReport || sheetPanelOpen

  const report = useMemo(() => buildReportModel(bundle, filters, new Date()), [bundle, filters])
  const issueExportRows = useMemo(
    () => selectIssueDetailExportRows(report.issueTable, ISSUE_ROWS_PER_EXPORT_SLIDE),
    [report.issueTable],
  )
  const issueExportPages = useMemo(
    () => chunkRows(issueExportRows, ISSUE_ROWS_PER_EXPORT_SLIDE),
    [issueExportRows],
  )
  const issueAppRows = useMemo(
    () => selectIssueDetailExportRows(report.issueTable, ISSUE_ROWS_PER_APP_SLIDE),
    [report.issueTable],
  )
  const issueAppPages = useMemo(
    () => chunkRows(issueAppRows, ISSUE_ROWS_PER_APP_SLIDE),
    [issueAppRows],
  )

  useEffect(() => {
    setIssuePageIndex(0)
  }, [issueAppRows])

  useEffect(() => {
    saveFilters(filters)
  }, [filters])

  useEffect(() => {
    clearLegacyConnectionData()
  }, [])

  useEffect(() => {
    if (!import.meta.env.PROD) return
    let active = true
    const poll = async (): Promise<void> => {
      const info = await checkForUpdate()
      if (!active) return
      setUpdateInfo(info)
      setLastUpdateCheck(new Date())
      if (info && dismissedUpdateVersionRef.current !== info.version) {
        setUpdateDefaultTab('update')
        setUpdateOpen(true)
      }
    }
    void poll()
    const interval = window.setInterval(() => void poll(), 15 * 60 * 1000)
    const handleVisibility = (): void => {
      if (document.visibilityState === 'visible') void poll()
    }
    document.addEventListener('visibilitychange', handleVisibility)
    return () => {
      active = false
      window.clearInterval(interval)
      document.removeEventListener('visibilitychange', handleVisibility)
    }
  }, [])

  async function handleManualUpdateCheck(): Promise<void> {
    if (updateChecking) return
    setUpdateChecking(true)
    try {
      const info = await checkForUpdate()
      setUpdateInfo(info)
      setLastUpdateCheck(new Date())
      if (info) {
        dismissedUpdateVersionRef.current = null
        setUpdateDefaultTab('update')
      }
    } finally {
      setUpdateChecking(false)
    }
  }

  function bundleForImports(next: Partial<Record<SheetRole, ImportedSheetFile>>): SheetBundle {
    const complete = (Object.keys(ROLE_CONFIG) as SheetRole[]).every((role) => Boolean(next[role]))
    return {
      source: complete ? 'files' : 'empty',
      sheets: {
        bimIssues: next.bimIssues?.sheet ?? EMPTY_BUNDLE.sheets.bimIssues,
        mechanical: next.mechanical?.sheet ?? EMPTY_BUNDLE.sheets.mechanical,
        electrical: next.electrical?.sheet ?? EMPTY_BUNDLE.sheets.electrical,
        welding: next.welding?.sheet ?? EMPTY_BUNDLE.sheets.welding,
      },
    }
  }

  function commitImportedBundle(
    nextImports: Partial<Record<SheetRole, ImportedSheetFile>>,
    nextBundle: SheetBundle,
  ): void {
    setImports(nextImports)
    setBundle(nextBundle)
    setActiveSlide('overview')
    if (nextBundle.source === 'files') setSheetPanelOpen(false)
  }

  function applyQualityReview(): void {
    if (!dataQualityReview) return
    const correctedBundle = applyDataQualityEdits(
      dataQualityReview.bundle,
      dataQualityReview.findings,
      dataQualityReview.values,
    )
    const correctedImports = { ...dataQualityReview.imports }
    ;(Object.keys(ROLE_CONFIG) as SheetRole[]).forEach((role) => {
      const imported = correctedImports[role]
      if (imported) correctedImports[role] = { ...imported, sheet: correctedBundle.sheets[role] }
    })
    commitImportedBundle(correctedImports, correctedBundle)
    setDataQualityReview(null)
  }

  async function handleFiles(files: File[]): Promise<void> {
    if (files.length === 0) return
    setImporting(true)
    setError(null)
    try {
      const expandedFiles = await expandImportFiles(files)
      const results = await Promise.allSettled(expandedFiles.map(importSpreadsheet))
      const next = { ...imports }
      const errors: string[] = []
      results.forEach((result) => {
        if (result.status === 'fulfilled') next[result.value.role] = result.value
        else errors.push(result.reason instanceof Error ? result.reason.message : 'A spreadsheet could not be imported.')
      })
      const nextBundle = bundleForImports(next)
      if (nextBundle.source === 'files') {
        const audit = auditDataQuality(nextBundle)
        if (audit.findings.length > 0) {
          setDataQualityReview({
            bundle: nextBundle,
            imports: next,
            findings: audit.findings,
            values: audit.initialValues,
          })
        } else {
          commitImportedBundle(next, nextBundle)
        }
      } else {
        commitImportedBundle(next, nextBundle)
      }
      if (errors.length > 0) {
        setError(errors.join(' '))
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Unable to import the selected spreadsheets.')
    } finally {
      setImporting(false)
    }
  }

  async function handleDrop(dataTransfer: DataTransfer): Promise<void> {
    setImporting(true)
    setError(null)
    try {
      const droppedFiles = await filesFromDrop(dataTransfer)
      if (droppedFiles.length === 0) throw new Error('The dropped folder did not contain any files.')
      await handleFiles(droppedFiles)
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Unable to read the dropped folder.')
    } finally {
      setImporting(false)
    }
  }

  function removeImport(role: SheetRole): void {
    setDataQualityReview(null)
    const next = { ...imports }
    delete next[role]
    setImports(next)
    setBundle(bundleForImports(next))
    setSheetPanelOpen(true)
  }

  function clearImports(): void {
    setDataQualityReview(null)
    setImports({})
    setBundle(EMPTY_BUNDLE)
    setSheetPanelOpen(true)
    setError(null)
  }

  const slide = activeSlide === 'overview'
    ? <OverviewSlide report={report} />
    : activeSlide === 'issues'
      ? (
          <IssueTableSlide
            report={report}
            rows={issueAppPages[issuePageIndex] ?? []}
            summaryRows={issueAppRows}
            pageIndex={issuePageIndex}
            pageCount={issueAppPages.length}
            onPreviousPage={() => setIssuePageIndex((current) => Math.max(0, current - 1))}
            onNextPage={() => setIssuePageIndex((current) => Math.min(issueAppPages.length - 1, current + 1))}
          />
        )
      : <FieldSlide report={report} />

  return (
    <div className="app">
      <input
        ref={fileInputRef}
        className="file-input"
        type="file"
        accept=".zip,.xls,.xlsx,.csv"
        multiple
        onChange={(event) => {
          void handleFiles(Array.from(event.target.files ?? []))
          event.target.value = ''
        }}
      />
      <input
        ref={folderInputRef}
        className="file-input"
        type="file"
        multiple
        {...({ webkitdirectory: '' } as React.InputHTMLAttributes<HTMLInputElement>)}
        onChange={(event) => {
          void handleFiles(Array.from(event.target.files ?? []))
          event.target.value = ''
        }}
      />
      <header className="app-header">
        <div className="brand-block">
          <div className="brand-mark">
            <ShieldCheck size={22} />
          </div>
          <div>
            <h1>QA/QC Intelligence</h1>
            <p>
              {workspaceTab === 'manual'
                ? 'Reconcile LotusWorks issues from ACC'
                : hasReport
                  ? `Weekly report | ${report.periodLabel}`
                  : 'Import weekly exports to generate a report'}
            </p>
          </div>
        </div>
        <div className="header-pills">
          {workspaceTab === 'manual' ? (
            <>
              <span className="connected">
                <ShieldCheck size={14} />
                LotusWorks only
              </span>
              <span>
                <GitBranch size={14} />
                ID-safe merge
              </span>
            </>
          ) : (
            <>
              <span className={cx('live-pill', filesReady && 'connected')}>
                <FileSpreadsheet size={14} />
                {bundle.source === 'demo' ? 'Preview Layout' : `${importCount}/4 Files Ready`}
              </span>
              <span>
                <CalendarClock size={14} />
                Current {report.currentWeek.label}
              </span>
              <span>
                <Gauge size={14} />
                GC Safe Export
              </span>
            </>
          )}
        </div>
        <div className="header-actions">
          <button
            className={cx('icon-button labeled update-button', updateInfo && 'has-update')}
            type="button"
            onClick={() => {
              setUpdateDefaultTab(updateInfo ? 'update' : 'changelog')
              setUpdateOpen(true)
            }}
          >
            <Bell size={16} />
            Updates
            {updateInfo && <span className="update-badge">v{updateInfo.version}</span>}
          </button>
          {workspaceTab === 'report' && (
            <>
              {hasReport && (
                <button
                  className="icon-button labeled"
                  type="button"
                  onClick={() => setSheetPanelOpen((prev) => !prev)}
                  aria-pressed={sheetPanelOpen}
                >
                  {sheetPanelOpen ? <PanelLeftClose size={16} /> : <PanelLeftOpen size={16} />}
                  {sheetPanelOpen ? 'Hide Files' : 'Files'}
                </button>
              )}
              <button className="icon-button labeled" type="button" onClick={() => fileInputRef.current?.click()} disabled={importing}>
                {importing ? <RefreshCw size={16} className="spin" /> : <Archive size={16} />}
                Import ZIP
              </button>
              <button
                className="button primary"
                type="button"
                disabled={exporting || pdfExporting || !hasReport}
                onClick={async () => {
                  setExporting(true)
                  try {
                    await exportReportDeck(report)
                  } catch (err) {
                    setError(err instanceof Error ? `PowerPoint export failed: ${err.message}` : 'PowerPoint export failed.')
                  } finally {
                    await waitForExportCooldown()
                    setExporting(false)
                  }
                }}
              >
                {exporting ? <RefreshCw size={16} className="spin" /> : <Download size={16} />}
                Export PPTX
              </button>
              <button
                className="button primary pdf-button"
                type="button"
                disabled={exporting || pdfExporting || !hasReport}
                onClick={async () => {
                  setPdfExporting(true)
                  try {
                    await exportSlidesPdf(report)
                  } catch (err) {
                    setError(err instanceof Error ? `PDF export failed: ${err.message}` : 'PDF export failed.')
                  } finally {
                    await waitForExportCooldown()
                    setPdfExporting(false)
                  }
                }}
              >
                {pdfExporting ? <RefreshCw size={16} className="spin" /> : <Download size={16} />}
                Export PDF
              </button>
            </>
          )}
        </div>
      </header>

      <nav className="workspace-tabs" aria-label="QA/QC tools">
        <button className={cx(workspaceTab === 'report' && 'active')} type="button" onClick={() => setWorkspaceTab('report')}>
          <Gauge size={16} />
          Weekly Report
        </button>
        <button className={cx(workspaceTab === 'manual' && 'active')} type="button" onClick={() => setWorkspaceTab('manual')}>
          <GitBranch size={16} />
          Manual Issues Update
        </button>
      </nav>

      {workspaceTab === 'report' && hasReport && (
        <section className="filter-strip">
          <div className="report-mode-picker" role="group" aria-label="Reporting mode">
            {([
              ['oac', 'OAC'],
              ['justine', 'Justine'],
              ['manual', 'Manual'],
            ] as const).map(([mode, label]) => (
              <button
                className={cx('report-mode-option', filters.reportingMode === mode && 'active')}
                type="button"
                key={mode}
                aria-pressed={filters.reportingMode === mode}
                title={mode === 'oac' ? 'Previous completed work week' : mode === 'justine' ? 'Previous and current work weeks' : 'Choose work weeks'}
                onClick={() => setFilters((prev) => {
                  if (prev.reportingMode === 'oac') oacContractorsRef.current = prev.contractors
                  return {
                    ...prev,
                    reportingMode: mode,
                    oac: mode !== 'manual',
                    contractors: mode === 'justine'
                      ? []
                      : mode === 'oac'
                        ? oacContractorsRef.current
                        : prev.contractors,
                  }
                })}
              >
                {label}
              </button>
            ))}
          </div>
          <FilterMenu
            label="Work Week"
            icon={<CalendarClock size={14} />}
            options={report.filterOptions.workWeeks}
            selected={filters.workWeeks}
            disabled={filters.oac}
            onChange={(workWeeks) => setFilters((prev) => ({ ...prev, workWeeks }))}
          />
          <FilterMenu
            label="Discipline"
            icon={<SlidersHorizontal size={14} />}
            options={report.filterOptions.disciplines}
            selected={filters.disciplines}
            onChange={(disciplines) => setFilters((prev) => ({ ...prev, disciplines }))}
          />
          <FilterMenu
            label="Contractor"
            icon={<Filter size={14} />}
            options={report.filterOptions.contractors}
            selected={filters.contractors}
            onChange={(contractors) => setFilters((prev) => {
              if (prev.reportingMode === 'oac') oacContractorsRef.current = contractors
              return { ...prev, contractors }
            })}
          />
          <FilterMenu
            label="Subtype"
            icon={<AlertCircle size={14} />}
            options={report.filterOptions.subtypes}
            selected={filters.subtypes}
            onChange={(subtypes) => setFilters((prev) => ({ ...prev, subtypes }))}
          />
          <FilterMenu
            label="Status"
            icon={<CheckCircle2 size={14} />}
            options={report.filterOptions.statuses}
            selected={filters.statuses}
            onChange={(statuses) => setFilters((prev) => ({ ...prev, statuses }))}
          />
        </section>
      )}

      {workspaceTab === 'report' && error && (
        <div className="error-banner">
          <AlertCircle size={16} />
          <span>{error}</span>
          <button type="button" onClick={() => setError(null)}>Dismiss</button>
        </div>
      )}

      <main className={cx('workspace', workspaceTab === 'manual' && 'workspace-manual', hasReport && !showSheetPanel && 'workspace-expanded')}>
        {workspaceTab === 'manual' ? (
          <ManualIssuesUpdate />
        ) : (
          <>
            {showSheetPanel && (
              <ImportPanel
                imports={imports}
                onChoose={() => fileInputRef.current?.click()}
                onChooseFolder={() => folderInputRef.current?.click()}
                onRemove={removeImport}
                onClear={clearImports}
                importing={importing}
              />
            )}
            <section className="report-workspace">
              {hasReport ? (
                <>
                  <nav className="slide-tabs">
                    <button className={cx(activeSlide === 'overview' && 'active')} type="button" onClick={() => setActiveSlide('overview')}>
                      <Gauge size={15} />
                      Overview
                    </button>
                    <button className={cx(activeSlide === 'issues' && 'active')} type="button" onClick={() => setActiveSlide('issues')}>
                      <Table2 size={15} />
                      Issue Detail
                    </button>
                    <button className={cx(activeSlide === 'field' && 'active')} type="button" onClick={() => setActiveSlide('field')}>
                      <ClipboardCheck size={15} />
                      Inspections & Welding
                    </button>
                  </nav>
                  {slide}
                </>
              ) : (
                <ConnectFirst
                  imports={imports}
                  importing={importing}
                  onChoose={() => fileInputRef.current?.click()}
                  onChooseFolder={() => folderInputRef.current?.click()}
                  onDrop={(dataTransfer) => void handleDrop(dataTransfer)}
                  onPreview={() => {
                    setBundle(PREVIEW_BUNDLE)
                    setActiveSlide('overview')
                    setSheetPanelOpen(false)
                  }}
                />
              )}
            </section>
          </>
        )}
      </main>

      <UpdateModal
        open={updateOpen}
        onClose={() => {
          if (updateInfo) dismissedUpdateVersionRef.current = updateInfo.version
          setUpdateOpen(false)
        }}
        info={updateInfo}
        defaultTab={updateDefaultTab}
        checking={updateChecking}
        lastChecked={lastUpdateCheck}
        onCheck={handleManualUpdateCheck}
      />

      <DataQualityModal
        open={Boolean(dataQualityReview)}
        findings={dataQualityReview?.findings ?? []}
        values={dataQualityReview?.values ?? {}}
        onChange={(id, value) => setDataQualityReview((current) => current
          ? { ...current, values: { ...current.values, [id]: value } }
          : current)}
        onCancel={() => setDataQualityReview(null)}
        onApply={applyQualityReview}
      />

      {hasReport && (
        <div className="export-deck" aria-hidden="true">
          <OverviewSlide report={report} exportable />
          {issueExportPages.map((rows, index) => (
            <IssueTableSlide
              key={`issues-${index}`}
              report={report}
              rows={rows}
              summaryRows={issueExportRows}
              pageIndex={index}
              pageCount={issueExportPages.length}
              exportable
            />
          ))}
          <FieldSlide report={report} exportable />
        </div>
      )}
    </div>
  )
}
