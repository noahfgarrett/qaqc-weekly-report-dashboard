import { useEffect, useMemo, useRef, useState } from 'react'
import {
  Activity,
  AlertCircle,
  Archive,
  Bell,
  CalendarClock,
  CheckCircle2,
  ChevronDown,
  ClipboardCheck,
  Download,
  FileSpreadsheet,
  Filter,
  Flame,
  FolderOpen,
  Gauge,
  GitBranch,
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
import { buildReportModel, mergeFilters } from '@/calculations/report'
import { buildSampleBundle } from '@/data/sampleData'
import { CHANGELOG } from '@/data/changelog'
import { exportReportDeck } from '@/export/pptx'
import { exportSlidesPdf } from '@/export/pdf'
import { expandImportFiles, filesFromDrop, importSpreadsheet } from '@/services/fileImport'
import { clearLegacyConnectionData, loadFilters, saveFilters } from '@/services/storage'
import { checkForUpdate, downloadUpdate } from '@/services/updateChecker'
import type {
  AgingBucket,
  ElectricalPoint,
  ImportedSheetFile,
  IssueDetailRow,
  KpiMetric,
  MonthlyIssuePoint,
  ReportFilters,
  SheetBundle,
  SheetRecord,
  SheetRole,
  UpdateInfo,
  WeeklyIssuePoint,
  WeldingPoint,
} from '@/types'
import { compactNumber, deltaLabel, percent } from '@/utils/format'
import { compareWorkWeeks } from '@/utils/workWeeks'

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
const ISSUE_ROWS_PER_EXPORT_SLIDE = 16

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
}: {
  open: boolean
  title: string
  children: React.ReactNode
  onClose: () => void
  wide?: boolean
}) {
  if (!open) return null
  return (
    <div className="modal-backdrop" role="presentation" onMouseDown={onClose}>
      <div
        className={cx('modal', wide && 'modal-wide')}
        role="dialog"
        aria-modal="true"
        aria-label={title}
        onMouseDown={(event) => event.stopPropagation()}
      >
        <div className="modal-header">
          <h2>{title}</h2>
          <button className="icon-button" type="button" onClick={onClose} aria-label="Close">
            <X size={18} />
          </button>
        </div>
        {children}
      </div>
    </div>
  )
}

function UpdateModal({
  open,
  onClose,
  info,
}: {
  open: boolean
  onClose: () => void
  info: UpdateInfo | null
}) {
  const [tab, setTab] = useState<'update' | 'changelog'>(info ? 'update' : 'changelog')
  const [downloading, setDownloading] = useState(false)
  const [downloadMessage, setDownloadMessage] = useState<string | null>(null)
  const [downloadError, setDownloadError] = useState<string | null>(null)

  useEffect(() => {
    if (open) {
      setTab(info ? 'update' : 'changelog')
      setDownloadMessage(null)
      setDownloadError(null)
    }
  }, [open, info])

  return (
    <Modal open={open} title={info ? 'Update Available' : 'Changelog'} onClose={onClose} wide>
      <div className="update-tabs">
        {info && (
          <button className={cx(tab === 'update' && 'active')} type="button" onClick={() => setTab('update')}>
            Update
          </button>
        )}
        <button className={cx(tab === 'changelog' && 'active')} type="button" onClick={() => setTab('changelog')}>
          Changelog
        </button>
      </div>
      {tab === 'update' && info ? (
        <div className="update-card">
          <div className="version-hop">
            <span>v{__APP_VERSION__}</span>
            <GitBranch size={15} />
            <strong>v{info.version}</strong>
          </div>
          <div className="release-notes">
            {(info.releaseNotes || 'A new dashboard build is available.')
              .split('\n')
              .filter(Boolean)
              .slice(0, 9)
              .map((line) => (
                <p key={line}>{line.replace(/^[-#* ]+/, '')}</p>
              ))}
          </div>
          <p className={cx('download-hint', downloadError && 'error')}>
            {downloadError ?? downloadMessage ?? `${info.assetName} will download directly to your browser downloads folder.`}
          </p>
          <div className="modal-actions">
            <button className="button secondary" type="button" onClick={onClose}>
              Later
            </button>
            <button
              className="button primary"
              type="button"
              disabled={downloading}
              onClick={async () => {
                setDownloading(true)
                setDownloadError(null)
                setDownloadMessage(null)
                try {
                  await downloadUpdate(info)
                  setDownloadMessage(`${info.assetName} is downloading.`)
                } catch (err) {
                  setDownloadError(err instanceof Error ? err.message : 'Unable to download this update.')
                } finally {
                  setDownloading(false)
                }
              }}
            >
              {downloading ? <RefreshCw size={16} className="spin" /> : <Download size={16} />}
              Download HTML v{info.version}
            </button>
          </div>
        </div>
      ) : (
        <div className="changelog-list">
          {CHANGELOG.map((entry) => (
            <article className="changelog-entry" key={entry.version}>
              <div>
                <strong>v{entry.version}</strong>
                <span>{new Date(entry.date).toLocaleDateString()}</span>
              </div>
              <ul>
                {entry.notes.map((note) => (
                  <li key={note}>{note}</li>
                ))}
              </ul>
            </article>
          ))}
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

function KpiTile({ metric, hero }: { metric: KpiMetric; hero?: boolean }) {
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
        <span className={cx('kpi-delta', toneClass)}>
          {arrow && <span aria-hidden="true">{arrow}</span>}
          {deltaText}
        </span>
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
      <section className="kpi-group primary">
        <div className="kpi-group-head">
          <span>Project to date</span>
          <em>Through {report.reportWeek.label}</em>
        </div>
        <div className="kpi-group-body">
          {KPI_GROUPS.primary.map((id) => {
            const metric = byId.get(id)
            return metric ? <KpiTile key={id} metric={metric} /> : null
          })}
        </div>
      </section>
      <section className="kpi-group this-week">
        <div className="kpi-group-head">
          <span>{report.reportWeek.label}</span>
        </div>
        <div className="kpi-group-body">
          {KPI_GROUPS.week.map((id) => {
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
    ? 'OAC range'
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

function shortWorkWeek(label: string): string {
  return label.replace(/^WW(\d{2})['’]\d{2}(\d{2})$/i, "$1'$2")
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
        <text className="axis-title" x={pad.l} y={pad.t - 3}>ISSUES / WEEK</text>
        <text className="axis-title remaining-axis-label" x={width - pad.r} y={pad.t - 3} textAnchor="end">REMAINING OPEN</text>

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

function MonthlyChart({ data }: { data: MonthlyIssuePoint[] }) {
  const [hoveredIndex, setHoveredIndex] = useState<number | null>(null)
  const visible = data.slice(-10)
  const { max, ticks } = niceScale(Math.max(1, ...visible.flatMap((point) => [point.opened, point.closed])))
  const width = 700
  const height = 150
  const pad = { l: 44, r: 18, t: 20, b: 30 }
  const chartW = width - pad.l - pad.r
  const chartH = height - pad.t - pad.b
  const step = chartW / Math.max(1, visible.length - 1)
  const toY = (value: number) => pad.t + chartH - (value / max) * chartH
  const pointsFor = (key: 'opened' | 'closed') => visible.map((point, index) => ({
    x: pad.l + index * step,
    y: toY(point[key]),
  }))
  const openedPoints = pointsFor('opened')
  const closedPoints = pointsFor('closed')
  const hovered = hoveredIndex === null ? null : visible[hoveredIndex]
  return (
    <div className="chart-shell mini-chart-shell">
      <div className="chart-legend compact-legend">
        <span className="legend-opened">Opened</span>
        <span className="legend-closed">Closed</span>
        <span className="legend-gap">Open Gap</span>
      </div>
      <svg className="mini-chart" viewBox={`0 0 ${width} ${height}`} role="img" aria-label="Cumulative Opened vs Closed">
        <defs>
          <linearGradient id="monthly-gap" x1="0" x2="0" y1="0" y2="1">
            <stop offset="0%" stopColor="#2e5aac" stopOpacity="0.12" />
            <stop offset="100%" stopColor="#2e5aac" stopOpacity="0.02" />
          </linearGradient>
        </defs>
        {ticks.map((tick) => (
          <g key={tick}>
            <line x1={pad.l} x2={width - pad.r} y1={toY(tick)} y2={toY(tick)} className="grid-line" />
            <text className="tick-label" x={pad.l - 8} y={toY(tick) + 4} textAnchor="end">{chartValue(tick)}</text>
          </g>
        ))}
        {visible.map((point, index) => (
          <g key={point.month} onMouseEnter={() => setHoveredIndex(index)} onMouseLeave={() => setHoveredIndex(null)}>
            <rect className="chart-hit-area" x={pad.l + index * step - step / 2} y={pad.t} width={step} height={chartH} />
            <text x={pad.l + index * step} y={height - 12} textAnchor="middle">{point.month}</text>
          </g>
        ))}
        <path d={`${smoothPath(openedPoints)} ${[...closedPoints].reverse().map((point) => `L ${point.x} ${point.y}`).join(' ')} Z`} className="monthly-gap-area" />
        <path d={smoothPath(openedPoints)} className="opened-line" fill="none" />
        <path d={smoothPath(closedPoints)} className="closed-line" fill="none" />
        {openedPoints.map((point, index) => (
          <g key={`opened-${visible[index].month}`}>
            <circle cx={point.x} cy={point.y} r={index === hoveredIndex ? 4 : 2.8} className="opened-dot" />
            <text className="series-data-label opened-label" x={point.x} y={Math.max(12, point.y - 9)} textAnchor="middle">{visible[index].opened}</text>
          </g>
        ))}
        {closedPoints.map((point, index) => (
          <g key={`closed-${visible[index].month}`}>
            <circle cx={point.x} cy={point.y} r={index === hoveredIndex ? 4 : 2.8} className="closed-dot" />
            <text className="series-data-label closed-label" x={point.x} y={Math.min(height - pad.b - 4, point.y + 15)} textAnchor="middle">{visible[index].closed}</text>
          </g>
        ))}
      </svg>
      {hovered && (
        <ChartTooltip
          compact
          title={hovered.month}
          entries={[`Opened ${hovered.opened}`, `Closed ${hovered.closed}`, `Gap ${hovered.gap}`]}
          xPercent={visible.length > 1 ? ((hoveredIndex ?? 0) / (visible.length - 1)) * 100 : 50}
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
  const visible = data.slice(-24)
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
              {(index % 4 === 0 || index === visible.length - 1) && (
                <text className="week-axis-label" x={x} y={height - 22} textAnchor="middle">
                  {shortWorkWeek(point.workWeek)}
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

function WeldingChart({ data }: { data: WeldingPoint[] }) {
  const [hoveredIndex, setHoveredIndex] = useState<number | null>(null)
  const visible = data.slice(-24)
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
        <line className="signoff-baseline" x1={pad.l} x2={width - pad.r} y1={toRateY(10)} y2={toRateY(10)} />
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
              {(index % 4 === 0 || index === visible.length - 1) && (
                <text className="week-axis-label" x={pad.l + index * group + group / 2} y={height - 22} textAnchor="middle">
                  {shortWorkWeek(point.workWeek)}
                </text>
              )}
            </g>
          )
        })}
        <path className="signoff-rate-line" d={smoothPath(ratePoints)} fill="none" />
        {ratePoints.map(({ point, x, y }, index) => {
          const labelY = index % 2 === 0
            ? Math.max(pad.t + 10, y - 10)
            : Math.min(pad.t + chartH - 5, y + 17)
          return (
            <g key={`rate-${point.workWeek}`}>
              <circle className="signoff-rate-dot" cx={x} cy={y} r="3.5" />
              <text className="signoff-rate-label" x={x} y={labelY} textAnchor="middle">{percent(point.signoffRate, 0)}</text>
            </g>
          )
        })}
        {visible.map((point, index) => {
          if (point.issuesCreated <= 0) return null
          const x = pad.l + index * group + group / 2
          const totalY = toY(point.total)
          const issueY = Math.max(pad.t + 10, totalY - 18)
          return (
            <g className="issue-badge" key={`issue-${point.workWeek}`} transform={`translate(${x} ${issueY})`}>
              <circle r="8" />
              <text textAnchor="middle" y="3">{point.issuesCreated}</text>
            </g>
          )
        })}
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
}: {
  children: React.ReactNode
  title: string
  icon: React.ReactNode
  exportable?: boolean
  meta?: string
  hideHeader?: boolean
}) {
  return (
    <section
      className={cx('slide-frame', hideHeader && 'headerless')}
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
            {meta && <em className="slide-meta">{meta}</em>}
          </header>
        )}
        {children}
      </div>
    </section>
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
            <Activity size={16} />
            <h3>Cumulative Opened vs Closed</h3>
          </div>
          <MonthlyChart data={report.monthlyTrend} />
        </article>
        <article className="panel chart-panel">
          <div className="panel-title">
            <CalendarClock size={16} />
            <h3>Issue Aging</h3>
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
  pageIndex = 0,
  pageCount = 1,
}: {
  report: ReturnType<typeof buildReportModel>
  exportable?: boolean
  rows?: IssueDetailRow[]
  pageIndex?: number
  pageCount?: number
}) {
  const groupCounts = report.issueTable.reduce<Record<string, number>>((acc, row) => {
    acc[row.group] = (acc[row.group] ?? 0) + 1
    return acc
  }, {})
  const visibleRows = rows ?? report.issueTable
  return (
    <SlideShell
      title="BIM Issues Detail"
      icon={<Table2 size={18} />}
      exportable={exportable}
      meta={pageCount > 1 ? `Page ${pageIndex + 1} of ${pageCount} | ${report.issueTable.length} issues` : `${report.issueTable.length} issues`}
    >
      <div className="detail-metrics">
        {['Open Carryover', 'Closed in Report Week', 'Opened + Closed in Report Week', 'Closed This Week'].map((group) => (
          <div className="metric-chip" key={group}>
            <strong>{compactNumber(groupCounts[group] ?? 0)}</strong>
            <span>{group.replace(' in Report Week', '')}</span>
          </div>
        ))}
      </div>
      <div className="issue-table-wrap">
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
              <tr key={`${row.id}-${row.workWeekClosed}-${index}`}>
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
    ['Welds Checked', compactNumber(report.summary.weldsChecked), deltaLabel(report.summary.deltas.weldsChecked), Flame],
    ['Welds Signed', compactNumber(report.summary.weldsSigned), deltaLabel(report.summary.deltas.weldsSigned), CheckCircle2],
  ] as const
  return (
    <SlideShell title="Inspections & Welding Signoffs" icon={<ClipboardCheck size={18} />} exportable={exportable}>
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
              <small>{report.reportWeek.label}</small>
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
          <WeldingChart data={report.welding} />
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

export default function App() {
  const [bundle, setBundle] = useState<SheetBundle>(EMPTY_BUNDLE)
  const [filters, setFilters] = useState<ReportFilters>(() => mergeFilters(loadFilters()))
  const [activeSlide, setActiveSlide] = useState<'overview' | 'issues' | 'field'>('overview')
  const [imports, setImports] = useState<Partial<Record<SheetRole, ImportedSheetFile>>>({})
  const [importing, setImporting] = useState(false)
  const [exporting, setExporting] = useState(false)
  const [pdfExporting, setPdfExporting] = useState(false)
  const [sheetPanelOpen, setSheetPanelOpen] = useState(true)
  const [error, setError] = useState<string | null>(null)
  const [updateInfo, setUpdateInfo] = useState<UpdateInfo | null>(null)
  const [updateOpen, setUpdateOpen] = useState(false)
  const fileInputRef = useRef<HTMLInputElement>(null)
  const folderInputRef = useRef<HTMLInputElement>(null)
  const importCount = Object.keys(imports).length
  const filesReady = importCount === 4
  const hasReport = bundle.source === 'files' || bundle.source === 'demo'
  const showSheetPanel = !hasReport || sheetPanelOpen

  const report = useMemo(() => buildReportModel(bundle, filters, new Date()), [bundle, filters])
  const issueExportPages = useMemo(
    () => chunkRows(report.issueTable, ISSUE_ROWS_PER_EXPORT_SLIDE),
    [report.issueTable],
  )

  useEffect(() => {
    saveFilters(filters)
  }, [filters])

  useEffect(() => {
    clearLegacyConnectionData()
  }, [])

  useEffect(() => {
    if (!import.meta.env.PROD) return
    checkForUpdate().then((info) => {
      if (info) {
        setUpdateInfo(info)
        setUpdateOpen(true)
      }
    })
  }, [])

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
      setImports(next)
      setBundle(nextBundle)
      setActiveSlide('overview')
      if (nextBundle.source === 'files') setSheetPanelOpen(false)
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
    const next = { ...imports }
    delete next[role]
    setImports(next)
    setBundle(bundleForImports(next))
    setSheetPanelOpen(true)
  }

  function clearImports(): void {
    setImports({})
    setBundle(EMPTY_BUNDLE)
    setSheetPanelOpen(true)
    setError(null)
  }

  const slide = activeSlide === 'overview'
    ? <OverviewSlide report={report} />
    : activeSlide === 'issues'
      ? <IssueTableSlide report={report} />
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
            <p>{hasReport ? `Weekly report through ${report.reportWeek.label}` : 'Import weekly exports to generate a report'}</p>
          </div>
        </div>
        <div className="header-pills">
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
        </div>
        <div className="header-actions">
          <button className="icon-button labeled" type="button" onClick={() => setUpdateOpen(true)}>
            <Bell size={16} />
            Updates
          </button>
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
            disabled={exporting || !hasReport}
            onClick={async () => {
              setExporting(true)
              try {
                await exportReportDeck(report)
              } catch (err) {
                setError(err instanceof Error ? `PowerPoint export failed: ${err.message}` : 'PowerPoint export failed.')
              } finally {
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
            disabled={pdfExporting || !hasReport}
            onClick={async () => {
              setPdfExporting(true)
              try {
                await exportSlidesPdf(report)
              } catch (err) {
                setError(err instanceof Error ? `PDF export failed: ${err.message}` : 'PDF export failed.')
              } finally {
                setPdfExporting(false)
              }
            }}
          >
            {pdfExporting ? <RefreshCw size={16} className="spin" /> : <Download size={16} />}
            Export PDF
          </button>
        </div>
      </header>

      {hasReport && (
        <section className="filter-strip">
          <button
            className={cx('oac-toggle', filters.oac && 'on')}
            type="button"
            onClick={() => setFilters((prev) => ({ ...prev, oac: !prev.oac }))}
          >
            <span>OAC</span>
            <strong>{filters.oac ? `Through ${report.reportWeek.label}` : 'Manual'}</strong>
          </button>
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
            onChange={(contractors) => setFilters((prev) => ({ ...prev, contractors }))}
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

      {error && (
        <div className="error-banner">
          <AlertCircle size={16} />
          <span>{error}</span>
          <button type="button" onClick={() => setError(null)}>Dismiss</button>
        </div>
      )}

      <main className={cx('workspace', hasReport && !showSheetPanel && 'workspace-expanded')}>
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
      </main>

      <UpdateModal open={updateOpen} onClose={() => setUpdateOpen(false)} info={updateInfo} />

      {hasReport && (
        <div className="export-deck" aria-hidden="true">
          <OverviewSlide report={report} exportable />
          {issueExportPages.map((rows, index) => (
            <IssueTableSlide
              key={`issues-${index}`}
              report={report}
              rows={rows}
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
