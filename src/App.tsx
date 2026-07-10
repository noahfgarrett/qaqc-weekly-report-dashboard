import { useEffect, useMemo, useRef, useState } from 'react'
import {
  Activity,
  AlertCircle,
  Bell,
  Bolt,
  CalendarClock,
  CheckCircle2,
  ChevronDown,
  ClipboardCheck,
  Download,
  FileUp,
  FileSpreadsheet,
  Filter,
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
  Upload,
  Wrench,
  X,
} from 'lucide-react'
import { buildReportModel, mergeFilters } from '@/calculations/report'
import { buildSampleBundle } from '@/data/sampleData'
import { CHANGELOG } from '@/data/changelog'
import { exportReportDeck } from '@/export/pptx'
import { exportSlidesPdf } from '@/export/pdf'
import { importSpreadsheet } from '@/services/fileImport'
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

type PlotPoint = {
  x: number
  y: number
}

function niceMax(value: number, steps = 4): number {
  if (value <= steps) return steps
  const magnitude = 10 ** Math.floor(Math.log10(value))
  const normalized = value / magnitude
  const step = normalized <= 2 ? magnitude / 2 : normalized <= 5 ? magnitude : magnitude * 2
  return Math.ceil(value / step) * step
}

function chartTicks(max: number, steps = 4): number[] {
  return Array.from({ length: steps + 1 }, (_, index) => max - (max / steps) * index)
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
  const max = niceMax(Math.max(1, ...visible.flatMap((point) => [point.opened, point.closed, point.remainingOpen])))
  const width = 980
  const height = 370
  const pad = { l: 62, r: 28, t: 24, b: 102 }
  const chartW = width - pad.l - pad.r
  const chartH = height - pad.t - pad.b
  const group = chartW / Math.max(visible.length, 1)
  const toY = (value: number) => pad.t + chartH - (value / max) * chartH
  const points = visible.map((point, index) => ({
    x: pad.l + index * group + group / 2,
    y: toY(point.remainingOpen),
    point,
  }))
  const hovered = hoveredIndex === null ? null : visible[hoveredIndex]
  const hoveredPoint = hoveredIndex === null ? null : points[hoveredIndex]

  return (
    <div className="chart-shell">
      <div className="chart-topline">
        <div className="chart-legend">
          <span className="legend-opened">Opened</span>
          <span className="legend-closed">Closed</span>
          <span className="legend-remaining">Remaining Open</span>
        </div>
        <span className="chart-window">Last {visible.length} reporting weeks</span>
      </div>
      <svg viewBox={`0 0 ${width} ${height}`} className="chart-svg" role="img" aria-label="Issues by Work Week">
        <defs>
          <linearGradient id="remaining-area" x1="0" x2="0" y1="0" y2="1">
            <stop offset="0%" stopColor="#f59e0b" stopOpacity="0.2" />
            <stop offset="100%" stopColor="#f59e0b" stopOpacity="0" />
          </linearGradient>
        </defs>
        {chartTicks(max).map((tick) => (
          <g key={tick}>
            <line
            x1={pad.l}
            x2={width - pad.r}
            y1={toY(tick)}
            y2={toY(tick)}
            className="grid-line"
            />
            <text className="tick-label" x={pad.l - 10} y={toY(tick) + 4} textAnchor="end">{chartValue(tick)}</text>
          </g>
        ))}
        {visible.map((point, index) => {
          const baseX = pad.l + index * group + group * 0.2
          const openedH = (point.opened / max) * chartH
          const closedH = (point.closed / max) * chartH
          const isActive = index === hoveredIndex
          const isLatest = index === visible.length - 1
          return (
            <g key={point.workWeek} onMouseEnter={() => setHoveredIndex(index)} onMouseLeave={() => setHoveredIndex(null)}>
              {isLatest && <rect className="report-week-band" x={pad.l + index * group} y={pad.t} width={group} height={chartH} rx="8" />}
              <rect className="chart-hit-area" x={pad.l + index * group} y={pad.t} width={group} height={chartH} />
              <rect className="bar opened" x={baseX} y={pad.t + chartH - openedH} width={group * 0.23} height={openedH} rx="5" />
              <rect className="bar closed" x={baseX + group * 0.31} y={pad.t + chartH - closedH} width={group * 0.23} height={closedH} rx="5" />
              {(isActive || isLatest) && point.opened > 0 && <text className="bar-value" x={baseX + group * 0.115} y={toY(point.opened) - 7} textAnchor="middle">{point.opened}</text>}
              {(isActive || isLatest) && point.closed > 0 && <text className="bar-value" x={baseX + group * 0.425} y={toY(point.closed) - 7} textAnchor="middle">{point.closed}</text>}
              {(index % 3 === 0 || index === visible.length - 1) && (
              <text x={pad.l + index * group + group / 2} y={height - 54} textAnchor="end" transform={`rotate(-45 ${pad.l + index * group + group / 2} ${height - 54})`}>
                  {point.workWeek.replace('WW', '')}
                </text>
              )}
            </g>
          )
        })}
        <path d={areaPath(points, pad.t + chartH)} className="remaining-area" />
        <path d={smoothPath(points)} className="remaining-line" fill="none" />
        {points.map(({ x, y, point }) => (
          <circle key={point.workWeek} cx={x} cy={y} r={point === hovered ? 5 : 3.5} className="remaining-dot" />
        ))}
        {points.length > 0 && (
          <text className="series-end-label" x={points[points.length - 1].x - 8} y={Math.max(pad.t + 12, points[points.length - 1].y - 10)} textAnchor="end">
            {`Open ${points[points.length - 1].point.remainingOpen}`}
          </text>
        )}
        <text x={width / 2} y={height - 18} className="axis-title" textAnchor="middle">Work Week</text>
        <text x="17" y={pad.t + chartH / 2} transform={`rotate(-90 17 ${pad.t + chartH / 2})`} className="axis-title" textAnchor="middle">Issue Count</text>
      </svg>
      {hovered && hoveredPoint && (
        <ChartTooltip
          title={hovered.workWeek}
          entries={[`Opened ${hovered.opened}`, `Closed ${hovered.closed}`, `Remaining ${hovered.remainingOpen}`]}
          xPercent={(hoveredPoint.x / width) * 100}
        />
      )}
    </div>
  )
}

function MonthlyChart({ data }: { data: MonthlyIssuePoint[] }) {
  const [hoveredIndex, setHoveredIndex] = useState<number | null>(null)
  const visible = data.slice(-8)
  const max = niceMax(Math.max(1, ...visible.flatMap((point) => [point.opened, point.closed])))
  const width = 460
  const height = 230
  const pad = { l: 48, r: 22, t: 18, b: 42 }
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
            <stop offset="0%" stopColor="#38bdf8" stopOpacity="0.18" />
            <stop offset="100%" stopColor="#14b8a6" stopOpacity="0.04" />
          </linearGradient>
        </defs>
        {chartTicks(max).map((tick) => (
          <g key={tick}>
            <line x1={pad.l} x2={width - pad.r} y1={toY(tick)} y2={toY(tick)} className="grid-line" />
            <text className="tick-label" x={pad.l - 8} y={toY(tick) + 4} textAnchor="end">{chartValue(tick)}</text>
          </g>
        ))}
        {visible.map((point, index) => (
          <g key={point.month} onMouseEnter={() => setHoveredIndex(index)} onMouseLeave={() => setHoveredIndex(null)}>
            <rect className="chart-hit-area" x={pad.l + index * step - step / 2} y={pad.t} width={step} height={chartH} />
            {(index % 2 === 0 || index === visible.length - 1) && (
              <text x={pad.l + index * step} y={height - 12} textAnchor="middle">{point.month}</text>
            )}
          </g>
        ))}
        <path d={`${smoothPath(openedPoints)} ${[...closedPoints].reverse().map((point) => `L ${point.x} ${point.y}`).join(' ')} Z`} className="monthly-gap-area" />
        <path d={smoothPath(openedPoints)} className="opened-line" fill="none" />
        <path d={smoothPath(closedPoints)} className="closed-line" fill="none" />
        {openedPoints.map((point, index) => <circle key={`opened-${visible[index].month}`} cx={point.x} cy={point.y} r={index === hoveredIndex ? 4 : 2.5} className="opened-dot" />)}
        {closedPoints.map((point, index) => <circle key={`closed-${visible[index].month}`} cx={point.x} cy={point.y} r={index === hoveredIndex ? 4 : 2.5} className="closed-dot" />)}
        {visible.length > 0 && (
          <>
            <text className="series-end-label opened-label" x={openedPoints[openedPoints.length - 1].x - 5} y={openedPoints[openedPoints.length - 1].y - 8} textAnchor="end">{visible[visible.length - 1].opened}</text>
            <text className="series-end-label closed-label" x={closedPoints[closedPoints.length - 1].x - 5} y={closedPoints[closedPoints.length - 1].y + 14} textAnchor="end">{visible[visible.length - 1].closed}</text>
          </>
        )}
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
      <div className="aging-summary">
        <strong>{compactNumber(total)}</strong>
        <span>Issues measured by age</span>
      </div>
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
  const max = niceMax(Math.max(1, ...visible.map((point) => point.finals)))
  const width = 620
  const height = 390
  const pad = { l: 54, r: 28, t: 24, b: 110 }
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
            <stop offset="0%" stopColor="#14b8a6" stopOpacity="0.26" />
            <stop offset="100%" stopColor="#14b8a6" stopOpacity="0" />
          </linearGradient>
        </defs>
        {chartTicks(max).map((tick) => (
          <g key={tick}>
            <line x1={pad.l} x2={width - pad.r} y1={toY(tick)} y2={toY(tick)} className="grid-line" />
            <text className="tick-label" x={pad.l - 9} y={toY(tick) + 4} textAnchor="end">{chartValue(tick)}</text>
          </g>
        ))}
        <path d={areaPath(points, pad.t + chartH)} className="electrical-area" />
        <path d={smoothPath(points)} className="electrical-line" fill="none" />
        {visible.map((point, index) => {
          const { x, y } = points[index]
          const isLatest = index === visible.length - 1
          const isActive = index === hoveredIndex
          return (
            <g key={point.workWeek} onMouseEnter={() => setHoveredIndex(index)} onMouseLeave={() => setHoveredIndex(null)}>
              <rect className="chart-hit-area" x={x - step / 2} y={pad.t} width={step} height={chartH} />
              <circle cx={x} cy={y} r={isActive ? 5 : 3.5} className="electrical-dot" />
              {point.issuesFound > 0 && (
                <g className="issue-badge" transform={`translate(${x} ${Math.max(pad.t + 12, y - 18)})`}>
                  <circle r="8" />
                  <text textAnchor="middle" y="3">{point.issuesFound}</text>
                </g>
              )}
              {isLatest && <text className="series-end-label electrical-label" x={x - 7} y={Math.max(pad.t + 12, y - 10)} textAnchor="end">{point.finals}</text>}
              {(index % 4 === 0 || isLatest) && (
                <text x={x} y={height - 62} textAnchor="end" transform={`rotate(-45 ${x} ${height - 62})`}>
                  {point.workWeek.replace('WW', '')}
                </text>
              )}
            </g>
          )
        })}
        <text x="17" y={pad.t + chartH / 2} transform={`rotate(-90 17 ${pad.t + chartH / 2})`} className="axis-title" textAnchor="middle">Final Inspections</text>
        <text x={width / 2} y={height - 18} className="axis-title" textAnchor="middle">Work Week Observed</text>
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
  const max = niceMax(Math.max(1, ...visible.flatMap((point) => [point.signed, point.total])))
  const width = 620
  const height = 390
  const pad = { l: 54, r: 54, t: 24, b: 110 }
  const chartW = width - pad.l - pad.r
  const chartH = height - pad.t - pad.b
  const group = chartW / Math.max(visible.length, 1)
  const toY = (value: number) => pad.t + chartH - (value / max) * chartH
  const toPercentY = (value: number) => pad.t + chartH - (value / 100) * chartH
  const linePoints = visible.map((point, index) => ({
    x: pad.l + index * group + group / 2,
    y: toPercentY(point.signoffRate),
  }))
  const hovered = hoveredIndex === null ? null : visible[hoveredIndex]
  return (
    <div className="chart-shell field-chart-shell">
      <svg className="chart-svg" viewBox={`0 0 ${width} ${height}`} role="img" aria-label="Welding Signoffs by Work Week">
        {chartTicks(max).map((tick) => (
          <g key={tick}>
            <line x1={pad.l} x2={width - pad.r} y1={toY(tick)} y2={toY(tick)} className="grid-line" />
            <text className="tick-label" x={pad.l - 9} y={toY(tick) + 4} textAnchor="end">{chartValue(tick)}</text>
          </g>
        ))}
        {[0, 25, 50, 75, 100].map((tick) => <text key={tick} className="percent-tick" x={width - pad.r + 9} y={toPercentY(tick) + 4}>{tick}%</text>)}
        <line x1={pad.l} x2={width - pad.r} y1={toPercentY(10)} y2={toPercentY(10)} className="baseline" />
        <text x={width - pad.r - 4} y={toPercentY(10) - 7} className="baseline-label" textAnchor="end">10% target</text>
        {visible.map((point, index) => {
          const x = pad.l + index * group + group * 0.14
          const signedH = (point.signed / max) * chartH
          const totalH = (point.total / max) * chartH
          const isLatest = index === visible.length - 1
          const isActive = index === hoveredIndex
          return (
            <g key={point.workWeek} onMouseEnter={() => setHoveredIndex(index)} onMouseLeave={() => setHoveredIndex(null)}>
              <rect className="chart-hit-area" x={pad.l + index * group} y={pad.t} width={group} height={chartH} />
              <rect className="bar closed" x={x} y={pad.t + chartH - signedH} width={group * 0.29} height={signedH} rx="5" />
              <rect className="bar opened" x={x + group * 0.35} y={pad.t + chartH - totalH} width={group * 0.29} height={totalH} rx="5" />
              {(isActive || isLatest) && point.signed > 0 && <text className="bar-value" x={x + group * 0.145} y={toY(point.signed) - 7} textAnchor="middle">{point.signed}</text>}
              {(isActive || isLatest) && point.total > 0 && <text className="bar-value" x={x + group * 0.495} y={toY(point.total) - 7} textAnchor="middle">{point.total}</text>}
              {point.issuesCreated > 0 && (
                <g className="issue-badge" transform={`translate(${x + group * 0.49} ${Math.max(pad.t + 12, toY(point.total) - 16)})`}>
                  <circle r="8" />
                  <text textAnchor="middle" y="3">{point.issuesCreated}</text>
                </g>
              )}
              {(index % 4 === 0 || isLatest) && (
                <text x={pad.l + index * group + group / 2} y={height - 62} textAnchor="end" transform={`rotate(-45 ${pad.l + index * group + group / 2} ${height - 62})`}>
                  {point.workWeek.replace('WW', '')}
                </text>
              )}
            </g>
          )
        })}
        <path d={smoothPath(linePoints)} className="signoff-line" fill="none" />
        {linePoints.map((point, index) => <circle key={`signoff-${visible[index].workWeek}`} cx={point.x} cy={point.y} r={index === hoveredIndex ? 4.5 : 2.7} className="signoff-dot" />)}
        {visible.length > 0 && <text className="series-end-label signoff-label" x={linePoints[linePoints.length - 1].x - 7} y={Math.max(pad.t + 13, linePoints[linePoints.length - 1].y - 10)} textAnchor="end">{percent(visible[visible.length - 1].signoffRate, 0)}</text>}
        <text x="17" y={pad.t + chartH / 2} transform={`rotate(-90 17 ${pad.t + chartH / 2})`} className="axis-title" textAnchor="middle">Weld Count</text>
        <text x={width - 17} y={pad.t + chartH / 2} transform={`rotate(90 ${width - 17} ${pad.t + chartH / 2})`} className="axis-title" textAnchor="middle">Sign-off %</text>
        <text x={width / 2} y={height - 18} className="axis-title" textAnchor="middle">Weld Work Week</text>
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
        <span>{report.source === 'files' ? 'Imported Files' : 'Preview Layout'}</span>
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
  imports,
  onChoose,
  onFiles,
  onPreview,
  importing,
}: {
  imports: Partial<Record<SheetRole, ImportedSheetFile>>
  onChoose: () => void
  onFiles: (files: File[]) => void
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
        onFiles(Array.from(event.dataTransfer.files))
      }}
    >
      <div className="connect-first-mark">
        <FileUp size={34} />
      </div>
      <h2>Drop weekly Smartsheet exports</h2>
      <p>{importCount}/4 required files ready</p>
      <div className="connect-first-actions">
        <button className="button primary" type="button" onClick={onChoose} disabled={importing}>
          {importing ? <RefreshCw size={16} className="spin" /> : <Upload size={16} />}
          Choose Files
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
      <small>.xls, .xlsx, and .csv</small>
    </section>
  )
}

function ImportPanel({
  imports,
  onChoose,
  onRemove,
  onClear,
  importing,
}: {
  imports: Partial<Record<SheetRole, ImportedSheetFile>>
  onChoose: () => void
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
        {importing ? <RefreshCw size={16} className="spin" /> : <Upload size={16} />}
        Import Exports
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
  const importCount = Object.keys(imports).length
  const filesReady = importCount === 4
  const hasReport = bundle.source === 'files' || bundle.source === 'demo'
  const showSheetPanel = !hasReport || sheetPanelOpen

  const report = useMemo(() => buildReportModel(bundle, filters, new Date()), [bundle, filters])

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
      const results = await Promise.allSettled(files.map(importSpreadsheet))
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
        accept=".xls,.xlsx,.csv"
        multiple
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
            {importing ? <RefreshCw size={16} className="spin" /> : <Upload size={16} />}
            Import Files
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
          <ImportPanel
            imports={imports}
            onChoose={() => fileInputRef.current?.click()}
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
              onFiles={(files) => void handleFiles(files)}
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
          <IssueTableSlide report={report} exportable />
          <FieldSlide report={report} exportable />
        </div>
      )}
    </div>
  )
}
