import { useEffect, useMemo, useState } from 'react'
import {
  Activity,
  AlertCircle,
  Bell,
  Bolt,
  CalendarClock,
  CheckCircle2,
  ChevronDown,
  ClipboardCheck,
  Database,
  Download,
  FileSpreadsheet,
  Filter,
  Gauge,
  GitBranch,
  KeyRound,
  LineChart,
  Link2,
  LockKeyhole,
  PanelLeftClose,
  PanelLeftOpen,
  RefreshCw,
  Search,
  ShieldCheck,
  SlidersHorizontal,
  Table2,
  Unlink,
  Wrench,
  X,
} from 'lucide-react'
import { buildReportModel, mergeFilters } from '@/calculations/report'
import { buildSampleBundle } from '@/data/sampleData'
import { CHANGELOG } from '@/data/changelog'
import { exportReportDeck } from '@/export/pptx'
import { exportSlidesPdf } from '@/export/pdf'
import {
  autoMapSheets,
  DEFAULT_ROLE_NAMES,
  listSmartsheets,
  loadMappedSheets,
} from '@/services/smartsheet'
import {
  clearSmartsheetToken,
  loadFilters,
  loadSheetMapping,
  loadSmartsheetToken,
  loadTokenMeta,
  saveFilters,
  saveSheetMapping,
  saveSmartsheetToken,
} from '@/services/storage'
import { checkForUpdate, downloadUpdate } from '@/services/updateChecker'
import type {
  AgingBucket,
  ElectricalPoint,
  IssueDetailRow,
  KpiMetric,
  MonthlyIssuePoint,
  ReportFilters,
  SheetBundle,
  SheetRecord,
  SheetRole,
  SmartSheetSummary,
  UpdateInfo,
  WeeklyIssuePoint,
  WeldingPoint,
} from '@/types'
import { compactNumber, deltaLabel, percent } from '@/utils/format'
import { compareWorkWeeks } from '@/utils/workWeeks'

const ROLE_CONFIG: Record<SheetRole, { label: string; icon: typeof AlertCircle; color: string }> = {
  bimIssues: { label: 'BIM Issues Log', icon: AlertCircle, color: '#38bdf8' },
  mechanical: { label: 'Mechanical / Process', icon: ClipboardCheck, color: '#14b8a6' },
  electrical: { label: 'Electrical Inspections', icon: Bolt, color: '#f59e0b' },
  welding: { label: 'Welding Signoffs', icon: Wrench, color: '#22c55e' },
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

function ConnectModal({
  open,
  onClose,
  onConnect,
  savedLabel,
  loading,
}: {
  open: boolean
  onClose: () => void
  onConnect: (token: string, remember: boolean) => void
  savedLabel?: string
  loading: boolean
}) {
  const [token, setToken] = useState('')
  const [remember, setRemember] = useState(true)

  useEffect(() => {
    if (open) setToken('')
  }, [open])

  return (
    <Modal open={open} title="Smartsheet Connection" onClose={onClose}>
      <div className="connect-panel">
        <div className="secure-mark">
          <LockKeyhole size={18} />
          <span>{savedLabel ? `Saved auth: ${savedLabel}` : 'Personal access token'}</span>
        </div>
        <label className="field">
          <span>API token</span>
          <input
            value={token}
            onChange={(event) => setToken(event.target.value)}
            type="password"
            placeholder="Paste Smartsheet API token"
            autoFocus
          />
        </label>
        <label className="check-row">
          <input
            type="checkbox"
            checked={remember}
            onChange={(event) => setRemember(event.target.checked)}
          />
          <span>Remember this device</span>
        </label>
        <div className="modal-actions">
          <button className="button secondary" type="button" onClick={onClose}>
            Cancel
          </button>
          <button
            className="button primary"
            type="button"
            onClick={() => onConnect(token.trim(), remember)}
            disabled={!token.trim() || loading}
          >
            {loading ? <RefreshCw size={16} className="spin" /> : <Link2 size={16} />}
            Connect
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

function KpiCard({ metric }: { metric: KpiMetric }) {
  const Icon = metric.icon
  return (
    <article className={cx('kpi-card', `tone-${metric.tone}`)}>
      <div className="kpi-topline">
        <span className="kpi-icon">
          <Icon size={17} />
        </span>
        <span className="kpi-delta">{metric.deltaLabel}</span>
      </div>
      <strong>{metric.value}</strong>
      <span>{metric.label}</span>
      <div className="sparkline" aria-hidden="true">
        <i />
        <i />
        <i />
        <i />
      </div>
    </article>
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
  const display = disabled
    ? 'OAC range'
    : selected.length === 0
      ? 'All'
      : selected.length === 1
        ? selected[0]
        : `${selected.length} selected`

  return (
    <div className="filter-menu">
      <button className="filter-button" type="button" onClick={() => setOpen((value) => !value)} disabled={disabled}>
        {icon}
        <span>{label}</span>
        <strong>{display}</strong>
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

function RangeChart({ data }: { data: WeeklyIssuePoint[] }) {
  const [hovered, setHovered] = useState<WeeklyIssuePoint | null>(null)
  const visible = data.slice(-30)
  const maxBar = Math.max(1, ...visible.flatMap((point) => [point.opened, point.closed]))
  const maxLine = Math.max(1, ...visible.map((point) => point.remainingOpen))
  const width = 960
  const height = 340
  const pad = { l: 46, r: 28, t: 22, b: 94 }
  const chartW = width - pad.l - pad.r
  const chartH = height - pad.t - pad.b
  const group = chartW / visible.length
  const points = visible.map((point, index) => ({
    x: pad.l + index * group + group / 2,
    y: pad.t + chartH - (point.remainingOpen / maxLine) * chartH,
    point,
  }))
  const path = points.map((point, index) => `${index === 0 ? 'M' : 'L'} ${point.x} ${point.y}`).join(' ')

  return (
    <div className="chart-shell">
      <div className="chart-legend">
        <span className="legend-opened">Opened</span>
        <span className="legend-closed">Closed</span>
        <span className="legend-remaining">Remaining Open</span>
      </div>
      <svg viewBox={`0 0 ${width} ${height}`} className="chart-svg" role="img" aria-label="Issues by Work Week">
        {[0, 0.25, 0.5, 0.75, 1].map((tick) => (
          <line
            key={tick}
            x1={pad.l}
            x2={width - pad.r}
            y1={pad.t + chartH * tick}
            y2={pad.t + chartH * tick}
            className="grid-line"
          />
        ))}
        {visible.map((point, index) => {
          const baseX = pad.l + index * group + group * 0.2
          const openedH = (point.opened / maxBar) * chartH * 0.82
          const closedH = (point.closed / maxBar) * chartH * 0.82
          return (
            <g key={point.workWeek} onMouseEnter={() => setHovered(point)} onMouseLeave={() => setHovered(null)}>
              <rect className="bar opened" x={baseX} y={pad.t + chartH - openedH} width={group * 0.22} height={openedH} rx="5" />
              <rect className="bar closed" x={baseX + group * 0.28} y={pad.t + chartH - closedH} width={group * 0.22} height={closedH} rx="5" />
              {(index % 3 === 0 || index === visible.length - 1) && (
              <text x={pad.l + index * group + group / 2} y={height - 48} textAnchor="end" transform={`rotate(-45 ${pad.l + index * group + group / 2} ${height - 48})`}>
                  {point.workWeek.replace('WW', '')}
                </text>
              )}
            </g>
          )
        })}
        <path d={path} className="remaining-line" fill="none" />
        {points.map(({ x, y, point }) => (
          <circle key={point.workWeek} cx={x} cy={y} r="3.5" className="remaining-dot" onMouseEnter={() => setHovered(point)} onMouseLeave={() => setHovered(null)} />
        ))}
      </svg>
      {hovered && (
        <div className="chart-tooltip">
          <strong>{hovered.workWeek}</strong>
          <span>Opened {hovered.opened}</span>
          <span>Closed {hovered.closed}</span>
          <span>Remaining {hovered.remainingOpen}</span>
        </div>
      )}
    </div>
  )
}

function MonthlyChart({ data }: { data: MonthlyIssuePoint[] }) {
  const visible = data.slice(-9)
  const max = Math.max(1, ...visible.flatMap((point) => [point.opened, point.closed]))
  const width = 420
  const height = 190
  const pad = { l: 34, r: 18, t: 18, b: 34 }
  const chartW = width - pad.l - pad.r
  const chartH = height - pad.t - pad.b
  const step = chartW / Math.max(1, visible.length - 1)
  const pathFor = (key: 'opened' | 'closed') =>
    visible.map((point, index) => `${index === 0 ? 'M' : 'L'} ${pad.l + index * step} ${pad.t + chartH - (point[key] / max) * chartH}`).join(' ')
  return (
    <svg className="mini-chart" viewBox={`0 0 ${width} ${height}`} role="img" aria-label="Cumulative Opened vs Closed">
      <line x1={pad.l} x2={width - pad.r} y1={pad.t + chartH} y2={pad.t + chartH} className="grid-line" />
      <path d={pathFor('opened')} className="opened-line" fill="none" />
      <path d={pathFor('closed')} className="closed-line" fill="none" />
      {visible.map((point, index) => (
        <text key={point.month} x={pad.l + index * step} y={height - 10} textAnchor="middle">
          {index % 2 === 0 ? point.month : ''}
        </text>
      ))}
    </svg>
  )
}

function AgingChart({ data }: { data: AgingBucket[] }) {
  const max = Math.max(1, ...data.map((bucket) => bucket.count))
  return (
    <div className="aging-chart">
      {data.map((bucket) => (
        <div className="aging-row" key={bucket.label}>
          <span>{bucket.label}</span>
          <div>
            <i style={{ width: `${(bucket.count / max) * 100}%`, background: bucket.color }} />
          </div>
          <strong>{bucket.count}</strong>
        </div>
      ))}
    </div>
  )
}

function ElectricalChart({ data }: { data: ElectricalPoint[] }) {
  const visible = data.slice(-24)
  const max = Math.max(1, ...visible.map((point) => point.finals))
  const width = 560
  const height = 360
  const pad = { l: 42, r: 24, t: 22, b: 126 }
  const chartW = width - pad.l - pad.r
  const chartH = height - pad.t - pad.b
  const step = chartW / Math.max(1, visible.length - 1)
  const path = visible.map((point, index) => `${index === 0 ? 'M' : 'L'} ${pad.l + index * step} ${pad.t + chartH - (point.finals / max) * chartH}`).join(' ')
  return (
    <svg className="chart-svg" viewBox={`0 0 ${width} ${height}`} role="img" aria-label="Electrical Inspections by Work Week">
      <line x1={pad.l} x2={width - pad.r} y1={pad.t + chartH} y2={pad.t + chartH} className="grid-line" />
      <text x="18" y="180" transform="rotate(-90 18 180)" className="axis-title">Inspection Count</text>
      <text x="226" y={height - 18} className="axis-title">Work Week Observed</text>
      <path d={path} className="electrical-line" fill="none" />
      {visible.map((point, index) => {
        const x = pad.l + index * step
        const y = pad.t + chartH - (point.finals / max) * chartH
        return (
          <g key={point.workWeek}>
            <circle cx={x} cy={y} r="3.5" className="electrical-dot" />
            {point.issuesFound > 0 && <circle cx={x} cy={y - 12} r="5" className="issue-marker" />}
            {(index % 5 === 0 || index === visible.length - 1) && (
              <text x={x} y={height - 86} textAnchor="end" transform={`rotate(-45 ${x} ${height - 86})`}>
                {point.workWeek.replace('WW', '')}
              </text>
            )}
          </g>
        )
      })}
    </svg>
  )
}

function WeldingChart({ data }: { data: WeldingPoint[] }) {
  const visible = data.slice(-24)
  const max = Math.max(1, ...visible.flatMap((point) => [point.signed, point.total]))
  const width = 560
  const height = 360
  const pad = { l: 42, r: 42, t: 22, b: 126 }
  const chartW = width - pad.l - pad.r
  const chartH = height - pad.t - pad.b
  const group = chartW / visible.length
  const line = visible.map((point, index) => `${index === 0 ? 'M' : 'L'} ${pad.l + index * group + group / 2} ${pad.t + chartH - (point.signoffRate / 100) * chartH}`).join(' ')
  return (
    <svg className="chart-svg" viewBox={`0 0 ${width} ${height}`} role="img" aria-label="Welding Signoffs by Work Week">
      <line x1={pad.l} x2={width - pad.r} y1={pad.t + chartH} y2={pad.t + chartH} className="grid-line" />
      <line x1={pad.l} x2={width - pad.r} y1={pad.t + chartH - chartH * 0.1} y2={pad.t + chartH - chartH * 0.1} className="baseline" />
      <text x={width - pad.r - 58} y={pad.t + chartH - chartH * 0.1 - 6} className="axis-title">10% Baseline</text>
      <text x="18" y="180" transform="rotate(-90 18 180)" className="axis-title">Weld Count</text>
      <text x={width - 16} y="180" transform={`rotate(90 ${width - 16} 180)`} className="axis-title">Sign-off %</text>
      {visible.map((point, index) => {
        const x = pad.l + index * group + group * 0.16
        const signedH = (point.signed / max) * chartH * 0.82
        const totalH = (point.total / max) * chartH * 0.82
        return (
          <g key={point.workWeek}>
            <rect className="bar closed" x={x} y={pad.t + chartH - signedH} width={group * 0.25} height={signedH} rx="5" />
            <rect className="bar opened" x={x + group * 0.3} y={pad.t + chartH - totalH} width={group * 0.25} height={totalH} rx="5" />
            {point.issuesCreated > 0 && <circle cx={x + group * 0.28} cy={pad.t + chartH - totalH - 10} r="5" className="issue-marker" />}
            {(index % 5 === 0 || index === visible.length - 1) && (
              <text x={pad.l + index * group + group / 2} y={height - 86} textAnchor="end" transform={`rotate(-45 ${pad.l + index * group + group / 2} ${height - 86})`}>
                {point.workWeek.replace('WW', '')}
              </text>
            )}
          </g>
        )
      })}
      <path d={line} className="signoff-line" fill="none" />
    </svg>
  )
}

function SlideShell({
  children,
  title,
  icon,
  exportable,
}: {
  children: React.ReactNode
  title: string
  icon: React.ReactNode
  exportable?: boolean
}) {
  return (
    <section className="slide-frame" {...(exportable ? { 'data-export-slide': true } : {})}>
      <div className="gc-guard top" />
      <div className="gc-guard bottom" />
      <div className="slide-safe">
        <header className="slide-header">
          <div>
            <span>{icon}</span>
            <h2>{title}</h2>
          </div>
        </header>
        {children}
      </div>
    </section>
  )
}

function OverviewSlide({ report, exportable }: { report: ReturnType<typeof buildReportModel>; exportable?: boolean }) {
  return (
    <SlideShell title="Weekly QA/QC Report" icon={<Gauge size={18} />} exportable={exportable}>
      <div className="report-pills">
        <span>OAC Through {report.reportWeek.label}</span>
        <span>{report.source === 'smartsheet' ? 'Smartsheet Live' : 'Preview Layout'}</span>
        <span>Report Week {report.reportWeek.label}</span>
      </div>
      <div className="kpi-grid">
        {report.kpis.map((metric) => (
          <KpiCard key={metric.id} metric={metric} />
        ))}
      </div>
      <div className="overview-grid">
        <article className="panel chart-panel wide">
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

function IssueTableSlide({ report, exportable }: { report: ReturnType<typeof buildReportModel>; exportable?: boolean }) {
  const groupCounts = report.issueTable.reduce<Record<string, number>>((acc, row) => {
    acc[row.group] = (acc[row.group] ?? 0) + 1
    return acc
  }, {})
  return (
    <SlideShell title="BIM Issues Detail" icon={<Table2 size={18} />} exportable={exportable}>
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
            {report.issueTable.slice(0, 16).map((row) => (
              <tr key={`${row.id}-${row.workWeekClosed}`}>
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
    ['Welds Checked', compactNumber(report.summary.weldsChecked), deltaLabel(report.summary.deltas.weldsChecked), Wrench],
    ['Welds Signed', compactNumber(report.summary.weldsSigned), deltaLabel(report.summary.deltas.weldsSigned), CheckCircle2],
    ['Avg Sign-off %', percent(report.summary.avgSignoffRate, 1), deltaLabel(report.summary.deltas.avgSignoffRate), Gauge],
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
      </div>
      <div className="field-chart-grid">
        <article className="panel chart-panel">
          <div className="panel-title">
            <Bolt size={16} />
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
            <Wrench size={16} />
            <h3>Welding Signoffs by Work Week</h3>
          </div>
          <div className="chart-legend">
            <span className="legend-closed">Signed Welds</span>
            <span className="legend-opened">Total Welds</span>
            <span className="legend-amber">Sign-off %</span>
            <span className="legend-issue">Issue Created</span>
          </div>
          <WeldingChart data={report.welding} />
        </article>
      </div>
    </SlideShell>
  )
}

function ConnectFirst({
  onConnect,
  onPreview,
  loading,
}: {
  onConnect: () => void
  onPreview: () => void
  loading: boolean
}) {
  return (
    <section className="connect-first">
      <div className="connect-first-mark">
        <Database size={34} />
      </div>
      <h2>Connect Smartsheet to generate the weekly report</h2>
      <p>
        The released app now starts clean. Once your token is saved, it will reconnect on launch and load the dashboard automatically.
      </p>
      <div className="connect-first-actions">
        <button className="button primary" type="button" onClick={onConnect}>
          <KeyRound size={16} />
          Connect Smartsheet
        </button>
        <button className="button secondary" type="button" onClick={onPreview} disabled={loading}>
          <LineChart size={16} />
          Preview Layout
        </button>
      </div>
      <div className="connect-first-grid">
        <span>BIM Issues Log</span>
        <span>Mechanical / Process Inspection Log</span>
        <span>Electrical Inspection Log</span>
        <span>Welding Signoffs</span>
      </div>
    </section>
  )
}

function SheetPanel({
  sheets,
  mapping,
  search,
  setSearch,
  onRoleChange,
  connected,
}: {
  sheets: SmartSheetSummary[]
  mapping: Partial<Record<SheetRole, number>>
  search: string
  setSearch: (value: string) => void
  onRoleChange: (role: SheetRole, sheetId: number) => void
  connected: boolean
}) {
  const visible = sheets.filter((sheet) => sheet.name.toLowerCase().includes(search.toLowerCase())).slice(0, 60)
  return (
    <aside className="sheet-panel">
      <div className="panel-title">
        <Database size={16} />
        <h3>Smartsheet Files</h3>
      </div>
      <div className="search-box">
        <Search size={15} />
        <input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Search sheets" />
      </div>
      <div className="role-map">
        {(Object.keys(ROLE_CONFIG) as SheetRole[]).map((role) => {
          const config = ROLE_CONFIG[role]
          const Icon = config.icon
          return (
            <label className="role-row" key={role}>
              <span style={{ color: config.color }}>
                <Icon size={15} />
              </span>
              <span>{config.label}</span>
              <select
                value={mapping[role] ?? ''}
                disabled={!connected}
                onChange={(event) => onRoleChange(role, Number(event.target.value))}
              >
                <option value="">Not linked</option>
                {sheets.map((sheet) => (
                  <option key={sheet.id} value={sheet.id}>
                    {sheet.name}
                  </option>
                ))}
              </select>
            </label>
          )
        })}
      </div>
      <div className="sheet-list">
        {visible.length === 0 ? (
          <div className="empty-sheet-list">
            <FileSpreadsheet size={18} />
            <span>{connected ? 'No sheets found' : 'Connect Smartsheet'}</span>
          </div>
        ) : (
          visible.map((sheet) => {
            const linked = Object.values(mapping).includes(sheet.id)
            return (
              <div className="sheet-row" key={sheet.id}>
                <FileSpreadsheet size={15} />
                <span>{sheet.name}</span>
                {linked && <CheckCircle2 size={14} />}
              </div>
            )
          })
        )}
      </div>
    </aside>
  )
}

export default function App() {
  const [bundle, setBundle] = useState<SheetBundle>(EMPTY_BUNDLE)
  const [filters, setFilters] = useState<ReportFilters>(() => mergeFilters(loadFilters()))
  const [activeSlide, setActiveSlide] = useState<'overview' | 'issues' | 'field'>('overview')
  const [connectOpen, setConnectOpen] = useState(false)
  const [loading, setLoading] = useState(false)
  const [exporting, setExporting] = useState(false)
  const [pdfExporting, setPdfExporting] = useState(false)
  const [sheetPanelOpen, setSheetPanelOpen] = useState(true)
  const [error, setError] = useState<string | null>(null)
  const [sheets, setSheets] = useState<SmartSheetSummary[]>([])
  const [sheetSearch, setSheetSearch] = useState('')
  const [mapping, setMapping] = useState<Partial<Record<SheetRole, number>>>(() => loadSheetMapping())
  const [updateInfo, setUpdateInfo] = useState<UpdateInfo | null>(null)
  const [updateOpen, setUpdateOpen] = useState(false)
  const tokenMeta = loadTokenMeta()
  const savedToken = loadSmartsheetToken()
  const connected = bundle.source === 'smartsheet'
  const hasReport = bundle.source === 'smartsheet' || bundle.source === 'demo'
  const showSheetPanel = !hasReport || sheetPanelOpen

  const report = useMemo(() => buildReportModel(bundle, filters, new Date()), [bundle, filters])

  useEffect(() => {
    saveFilters(filters)
  }, [filters])

  useEffect(() => {
    if (!import.meta.env.PROD) return
    checkForUpdate().then((info) => {
      if (info) {
        setUpdateInfo(info)
        setUpdateOpen(true)
      }
    })
  }, [])

  async function connect(token: string, remember: boolean): Promise<void> {
    setLoading(true)
    setError(null)
    try {
      const sheetList = await listSmartsheets(token)
      const autoMapping = { ...autoMapSheets(sheetList), ...loadSheetMapping() }
      const loaded = await loadMappedSheets(token, autoMapping)
      const nextBundle: SheetBundle = {
        source: 'smartsheet',
        sheets: {
          ...EMPTY_BUNDLE.sheets,
          ...loaded,
        },
      }
      if (remember) saveSmartsheetToken(token)
      setSheets(sheetList)
      setMapping(autoMapping)
      saveSheetMapping(autoMapping)
      setBundle(nextBundle)
      setSheetPanelOpen(false)
      setConnectOpen(false)
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Unable to connect to Smartsheet.')
    } finally {
      setLoading(false)
    }
  }

  async function refreshFromSaved(): Promise<void> {
    const token = loadSmartsheetToken()
    if (!token) {
      setConnectOpen(true)
      return
    }
    await connect(token, true)
  }

  useEffect(() => {
    if (savedToken) void refreshFromSaved()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  async function handleRoleChange(role: SheetRole, sheetId: number): Promise<void> {
    const token = loadSmartsheetToken()
    const next = { ...mapping, [role]: sheetId }
    setMapping(next)
    saveSheetMapping(next)
    if (!token) return
    setLoading(true)
    try {
      const loaded = await loadMappedSheets(token, next)
      setBundle({
        source: 'smartsheet',
        sheets: { ...EMPTY_BUNDLE.sheets, ...loaded },
      })
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Unable to load mapped sheet.')
    } finally {
      setLoading(false)
    }
  }

  function clearAuth(): void {
    clearSmartsheetToken()
    setBundle(EMPTY_BUNDLE)
    setSheets([])
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
      <header className="app-header">
        <div className="brand-block">
          <div className="brand-mark">
            <ShieldCheck size={22} />
          </div>
          <div>
            <h1>QA/QC Intelligence</h1>
            <p>{hasReport ? `Weekly report through ${report.reportWeek.label}` : 'Connect Smartsheet to generate a report'}</p>
          </div>
        </div>
        <div className="header-pills">
          <span className={cx('live-pill', connected && 'connected')}>
            <Database size={14} />
            {connected ? 'Smartsheet Live' : bundle.source === 'demo' ? 'Preview Layout' : 'Not Connected'}
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
              {sheetPanelOpen ? 'Hide Sheets' : 'Sheets'}
            </button>
          )}
          <button className="icon-button labeled" type="button" onClick={() => setConnectOpen(true)}>
            <KeyRound size={16} />
            {connected ? 'Auth' : 'Connect'}
          </button>
          <button className="icon-button labeled" type="button" onClick={refreshFromSaved} disabled={loading}>
            <RefreshCw size={16} className={cx(loading && 'spin')} />
            Refresh
          </button>
          <button
            className="button primary"
            type="button"
            disabled={exporting || !hasReport}
            onClick={async () => {
              setExporting(true)
              try {
                await exportReportDeck(report)
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
          <SheetPanel
            sheets={sheets}
            mapping={mapping}
            search={sheetSearch}
            setSearch={setSheetSearch}
            onRoleChange={(role, sheetId) => void handleRoleChange(role, sheetId)}
            connected={connected}
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
              loading={loading}
              onConnect={() => setConnectOpen(true)}
              onPreview={() => {
                setBundle(PREVIEW_BUNDLE)
                setActiveSlide('overview')
                setSheetPanelOpen(false)
              }}
            />
          )}
        </section>
      </main>

      <ConnectModal
        open={connectOpen}
        onClose={() => setConnectOpen(false)}
        onConnect={(token, remember) => void connect(token, remember)}
        savedLabel={tokenMeta?.masked}
        loading={loading}
      />
      <UpdateModal open={updateOpen} onClose={() => setUpdateOpen(false)} info={updateInfo} />

      {hasReport && (
        <div className="export-deck" aria-hidden="true">
          <OverviewSlide report={report} exportable />
          <IssueTableSlide report={report} exportable />
          <FieldSlide report={report} exportable />
        </div>
      )}

      {savedToken && (
        <button className="forget-token" type="button" onClick={clearAuth}>
          <Unlink size={14} />
          Forget token
        </button>
      )}
    </div>
  )
}
