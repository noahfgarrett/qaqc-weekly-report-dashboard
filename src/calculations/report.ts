import {
  AlertCircle,
  CheckCircle2,
  ClipboardCheck,
  ClipboardList,
  Gauge,
  TrendingDown,
  TrendingUp,
  Wrench,
} from 'lucide-react'
import type {
  AgingBucket,
  ElectricalPoint,
  FilterOptions,
  IssueDetailRow,
  KpiMetric,
  ReportFilters,
  ReportModel,
  SheetBundle,
  SheetRole,
  WeeklyIssuePoint,
  WeldingPoint,
  WorkWeek,
} from '@/types'
import { compactNumber, deltaLabel, percent } from '@/utils/format'
import {
  compareWorkWeeks,
  daysBetween,
  enumerateWorkWeeks,
  formatDate,
  parseDate,
  parseWorkWeek,
  previousWorkWeek,
  toIsoWorkWeek,
  workWeekEnd,
} from '@/utils/workWeeks'

const START_WEEK = 'WW51\'2025'

const EMPTY_FILTERS: ReportFilters = {
  oac: true,
  reportingMode: 'oac',
  workWeeks: [],
  disciplines: [],
  contractors: [],
  subtypes: [],
  statuses: [],
}

type IssueStatus = 'open' | 'pending' | 'closed' | 'void'

interface IssueRecord {
  id: string
  status: string
  statusKind: IssueStatus
  subtype: string
  title: string
  contractor: string
  discipline: string
  createdOn: Date | null
  updatedOn: Date | null
  dueDate: Date | null
  createdWeek: WorkWeek | null
  closedWeek: WorkWeek | null
  raw: Record<string, unknown>
}

interface InspectionRecord {
  phase: string
  workWeek: WorkWeek | null
  issue: string
  contractor: string
  discipline: string
  subtype: string
}

interface WeldRecord {
  no: string
  workWeek: WorkWeek | null
  signature: string
  issueCreated: string
  contractor: string
  discipline: string
  subtype: string
}

function value(row: Record<string, unknown>, keys: string[]): string {
  const normalized = new Map(
    Object.entries(row).map(([key, val]) => [key.toLowerCase().replace(/[^a-z0-9]/g, ''), val]),
  )
  for (const key of keys) {
    const direct = row[key]
    if (direct !== undefined && direct !== null) return String(direct)
    const loose = normalized.get(key.toLowerCase().replace(/[^a-z0-9]/g, ''))
    if (loose !== undefined && loose !== null) return String(loose)
  }
  return ''
}

function hasColumn(row: Record<string, unknown>, keys: string[]): boolean {
  const headers = new Set(Object.keys(row).map((key) => key.toLowerCase().replace(/[^a-z0-9]/g, '')))
  return keys.some((key) => headers.has(key.toLowerCase().replace(/[^a-z0-9]/g, '')))
}

function normalizeStatus(status: string): IssueStatus {
  const lower = status.trim().toLowerCase()
  if (lower === 'void') return 'void'
  if (lower === 'closed') return 'closed'
  if (lower === 'pending') return 'pending'
  return 'open'
}

function legacyIssueValue(row: Record<string, unknown>, field: 'createdBy' | 'createdOn' | 'closedOn'): string {
  if (field === 'createdBy') return value(row, ['BIM360_Created By', 'BIM360 Created By'])
  if (field === 'createdOn') return value(row, ['BIM360_Created On', 'BIM360 Created On'])
  return value(row, ['BIM360_Closed On', 'BIM360 Closed On'])
}

function issueContractor(id: string, contractor: string): string {
  const populated = contractor.trim()
  if (populated) return populated
  return id.match(/\d+$/)?.[0] === '1021' ? 'Bechtel' : 'Unassigned'
}

function shouldIncludeIssue(row: Record<string, unknown>): boolean {
  const legacyCreatedBy = legacyIssueValue(row, 'createdBy').trim()
  const hasLegacyData = Boolean(
    legacyCreatedBy
    || legacyIssueValue(row, 'createdOn').trim()
    || legacyIssueValue(row, 'closedOn').trim(),
  )
  if (!hasLegacyData) {
    const standardCreatorAliases = ['Created By', 'Issue Owner']
    if (!hasColumn(row, standardCreatorAliases)) return true
    return value(row, standardCreatorAliases).toLowerCase().includes('lotusworks')
  }
  return legacyCreatedBy.toLowerCase().includes('lotusworks')
}

function normalizeIssue(row: Record<string, unknown>): IssueRecord {
  const id = value(row, ['ID', 'Issue ID', 'BIM ID']) || String(row.__rowNumber ?? row.__rowId ?? '')
  const rawStatus = value(row, ['Status']).trim() || 'Open'
  const statusKind = normalizeStatus(rawStatus)
  const status = statusKind === 'pending' ? 'Pending' : rawStatus
  const legacyCreatedOn = legacyIssueValue(row, 'createdOn').trim()
  const legacyClosedOn = legacyIssueValue(row, 'closedOn').trim()
  const closedAt = value(row, ['Closed At']).trim()
  const createdOn = parseDate(legacyCreatedOn || value(row, ['Created On', 'Created', 'Date Created']))
  const updatedOn = parseDate(legacyClosedOn || closedAt || value(row, ['Updated On', 'Updated', 'Closed On', 'Date Closed']))
  const accType = value(row, ['Type']).trim()
  return {
    id,
    status,
    statusKind,
    subtype: accType || value(row, ['Subtype', 'Sub Type', 'Issue Subtype']) || 'Uncategorized',
    title: value(row, ['Title', 'Issue', 'Description']) || 'Untitled issue',
    contractor: issueContractor(id, value(row, ['Contractor', 'Responsible Contractor'])),
    discipline: value(row, ['Discipline', 'Trade']) || 'Unassigned',
    createdOn,
    updatedOn,
    dueDate: parseDate(value(row, ['Due Date', 'Due'])),
    createdWeek: createdOn ? toIsoWorkWeek(createdOn) : null,
    closedWeek: statusKind === 'closed' && updatedOn ? toIsoWorkWeek(updatedOn) : null,
    raw: row,
  }
}

function normalizeInspection(row: Record<string, unknown>): InspectionRecord {
  return {
    phase: value(row, ['Inspection Phase', 'Phase of Inspection', 'Inspection Status', 'Phase']),
    workWeek: parseWorkWeek(value(row, ['Work Week Observed', 'Observed Work Week', 'Week Observed', 'WW Observed', 'Work Week', 'WW'])),
    issue: value(row, ['Issue?', 'Issues?', 'Issue Found?', 'Issues Found', 'Issue']),
    contractor: value(row, ['General Contractor'])
      || value(row, ['Contractor', 'Responsible Contractor'])
      || 'Unassigned',
    discipline: value(row, ['Discipline', 'Trade']) || 'Unassigned',
    subtype: value(row, ['Subtype', 'Sub Type', 'Inspection Type']) || 'Uncategorized',
  }
}

function normalizeWeld(row: Record<string, unknown>): WeldRecord {
  return {
    no: value(row, ['NO', 'No.', 'Weld No', 'Weld Number', 'Weld #', 'Weld ID']),
    workWeek: parseWorkWeek(value(row, ['WELD WORK WEEK', 'Weld Work Week', 'Weld WW', 'Work Week Welded', 'Work Week'])),
    signature: value(row, ['SIGNATURE', 'Weld Signature', 'Inspector Signature', 'Signed By', 'Signed']),
    issueCreated: value(row, ['ISSUE CREATED? (Put BIM # If Yes)', 'Issue Created?', 'Issue Created', 'BIM #']),
    contractor: value(row, ['Contractor', 'Responsible Contractor']) || 'Unassigned',
    discipline: value(row, ['Discipline', 'Trade']) || 'Welding',
    subtype: value(row, ['Subtype', 'Sub Type', 'Weld Type']) || 'Welding',
  }
}

function selected(valueToTest: string, selectedValues: string[]): boolean {
  return selectedValues.length === 0 || selectedValues.includes(valueToTest)
}

function passesIssueFilters(issue: IssueRecord, filters: ReportFilters, reportWeek: WorkWeek): boolean {
  if (issue.statusKind === 'void') return false
  const week = issue.createdWeek?.label ?? ''
  if (filters.oac) {
    if (!issue.createdWeek || compareWorkWeeks(issue.createdWeek, reportWeek) > 0) return false
  } else if (!selected(week, filters.workWeeks)) {
    return false
  }
  return selected(issue.discipline, filters.disciplines)
    && selected(issue.contractor, filters.contractors)
    && selected(issue.subtype, filters.subtypes)
    && selected(issue.status, filters.statuses)
}

function passesInspectionFilters(row: InspectionRecord, filters: ReportFilters, reportWeek: WorkWeek): boolean {
  if (filters.oac) {
    if (!row.workWeek || compareWorkWeeks(row.workWeek, reportWeek) > 0) return false
  } else if (!selected(row.workWeek?.label ?? '', filters.workWeeks)) {
    return false
  }
  return selected(row.discipline, filters.disciplines)
    && selected(row.contractor, filters.contractors)
    && selected(row.subtype, filters.subtypes)
}

function passesWeldFilters(row: WeldRecord, filters: ReportFilters, reportWeek: WorkWeek): boolean {
  if (filters.oac) {
    if (!row.workWeek || compareWorkWeeks(row.workWeek, reportWeek) > 0) return false
  } else if (!selected(row.workWeek?.label ?? '', filters.workWeeks)) {
    return false
  }
  return selected(row.discipline, filters.disciplines)
    && selected(row.subtype, filters.subtypes)
}

function uniqueSorted(values: string[]): string[] {
  return [...new Set(values.filter(Boolean))].sort((a, b) => a.localeCompare(b))
}

function buildFilterOptions(
  issues: IssueRecord[],
  inspections: InspectionRecord[],
  welds: WeldRecord[],
): FilterOptions {
  const workWeeks = uniqueSorted([
    ...issues.map((issue) => issue.createdWeek?.label ?? ''),
    ...inspections.map((row) => row.workWeek?.label ?? ''),
    ...welds.map((row) => row.workWeek?.label ?? ''),
  ]).sort(compareWorkWeeks)

  return {
    workWeeks,
    disciplines: uniqueSorted([
      ...issues.map((issue) => issue.discipline),
      ...inspections.map((row) => row.discipline),
      ...welds.map((row) => row.discipline),
    ]),
    contractors: uniqueSorted([
      ...issues.map((issue) => issue.contractor),
      ...inspections.map((row) => row.contractor),
      ...welds.map((row) => row.contractor),
    ]),
    subtypes: uniqueSorted([
      ...issues.map((issue) => issue.subtype),
      ...inspections.map((row) => row.subtype),
      ...welds.map((row) => row.subtype),
    ]),
    statuses: uniqueSorted(issues.map((issue) => issue.status).filter((status) => normalizeStatus(status) !== 'void')),
  }
}

function countOpenedThrough(issues: IssueRecord[], week: WorkWeek): number {
  return issues.filter((issue) =>
    issue.statusKind !== 'void'
    && issue.createdWeek
    && compareWorkWeeks(issue.createdWeek, week) <= 0,
  ).length
}

function countClosedThrough(issues: IssueRecord[], week: WorkWeek): number {
  return issues.filter((issue) =>
    issue.statusKind === 'closed'
    && issue.closedWeek
    && compareWorkWeeks(issue.closedWeek, week) <= 0,
  ).length
}

function countOpenedIn(issues: IssueRecord[], week: WorkWeek): number {
  return issues.filter((issue) =>
    issue.statusKind !== 'void'
    && issue.createdWeek?.label === week.label,
  ).length
}

function countClosedIn(issues: IssueRecord[], week: WorkWeek): number {
  return issues.filter((issue) =>
    issue.statusKind === 'closed'
    && issue.closedWeek?.label === week.label,
  ).length
}

function isWeekInRange(week: WorkWeek | null, start: WorkWeek, end: WorkWeek): boolean {
  return Boolean(
    week
    && compareWorkWeeks(week, start) >= 0
    && compareWorkWeeks(week, end) <= 0,
  )
}

function countOpenedBetween(issues: IssueRecord[], start: WorkWeek, end: WorkWeek): number {
  return issues.filter((issue) => issue.statusKind !== 'void' && isWeekInRange(issue.createdWeek, start, end)).length
}

function countClosedBetween(issues: IssueRecord[], start: WorkWeek, end: WorkWeek): number {
  return issues.filter((issue) => issue.statusKind === 'closed' && isWeekInRange(issue.closedWeek, start, end)).length
}

function makeMetric(
  id: string,
  label: string,
  rawValue: number,
  previous: number,
  icon: KpiMetric['icon'],
  tone: KpiMetric['tone'] = 'neutral',
  formatter: (value: number) => string = compactNumber,
  spark?: number[],
): KpiMetric {
  const delta = rawValue - previous
  return {
    id,
    label,
    rawValue,
    value: formatter(rawValue),
    delta,
    deltaLabel: deltaLabel(delta),
    tone,
    icon,
    spark: spark && spark.length > 1 ? spark : undefined,
  }
}

function phaseEquals(phase: string, expected: string): boolean {
  const normalized = phase.trim().toLowerCase()
  const target = expected.toLowerCase()
  if (normalized === target) return true
  if (target === 'final') {
    return normalized.startsWith('final ') || normalized.includes('final inspection')
  }
  return false
}

function includesPhase(phase: string, expected: string): boolean {
  return phase.trim().toLowerCase().includes(expected.toLowerCase())
}

function buildWeeklyTrend(issues: IssueRecord[], reportWeek: WorkWeek): WeeklyIssuePoint[] {
  return enumerateWorkWeeks(START_WEEK, reportWeek).map((week) => ({
    workWeek: week.label,
    opened: countOpenedIn(issues, week),
    closed: countClosedIn(issues, week),
    remainingOpen: countOpenedThrough(issues, week) - countClosedThrough(issues, week),
  }))
}

function buildAging(issues: IssueRecord[], today: Date): AgingBucket[] {
  const buckets: AgingBucket[] = [
    { label: '0-14 Days', count: 0, color: '#B8C4D0' },
    { label: '14-28 Days', count: 0, color: '#778C9F' },
    { label: '28+ Days', count: 0, color: '#3F5367' },
  ]
  issues
    .filter((issue) => issue.statusKind !== 'void' && issue.createdOn)
    .forEach((issue) => {
      const end = issue.statusKind === 'closed' && issue.updatedOn ? issue.updatedOn : today
      const age = daysBetween(issue.createdOn as Date, end)
      if (age < 14) buckets[0].count += 1
      else if (age < 28) buckets[1].count += 1
      else buckets[2].count += 1
    })
  return buckets
}

function buildOpenAging(issues: IssueRecord[], today: Date): AgingBucket[] {
  const buckets: AgingBucket[] = [
    { label: '0-1', count: 0, color: '#B8C4D0' },
    { label: '1-2', count: 0, color: '#9EAFBF' },
    { label: '2-3', count: 0, color: '#778C9F' },
    { label: '3-4', count: 0, color: '#5D7185' },
    { label: '4+', count: 0, color: '#3F5367' },
  ]

  issues
    .filter((issue) => (issue.statusKind === 'open' || issue.statusKind === 'pending') && issue.createdOn)
    .forEach((issue) => {
      const age = daysBetween(issue.createdOn as Date, today)
      if (age < 7) buckets[0].count += 1
      else if (age < 14) buckets[1].count += 1
      else if (age < 21) buckets[2].count += 1
      else if (age < 28) buckets[3].count += 1
      else buckets[4].count += 1
    })

  return buckets
}

function buildIssueTable(
  issues: IssueRecord[],
  currentWeek: WorkWeek,
  periodStart: WorkWeek,
  periodEnd: WorkWeek,
  activityWindowOnly: boolean,
): IssueDetailRow[] {
  const reportRows: IssueDetailRow[] = []
  const discussionRows: IssueDetailRow[] = []

  issues.forEach((issue) => {
      if (issue.statusKind === 'void' || !issue.createdWeek) return
      const openedInReport = isWeekInRange(issue.createdWeek, periodStart, periodEnd)
      const closedInReport = isWeekInRange(issue.closedWeek, periodStart, periodEnd)
      const closedThisWeek = issue.closedWeek?.label === currentWeek.label
      const openedThisWeek = issue.createdWeek.label === currentWeek.label
      const createdThroughReport = compareWorkWeeks(issue.createdWeek, periodEnd) <= 0
      const isReportRow = activityWindowOnly
        ? openedInReport || closedInReport || closedThisWeek
        : issue.statusKind !== 'closed'
          ? createdThroughReport
          : closedInReport || (closedThisWeek && compareWorkWeeks(issue.createdWeek, currentWeek) < 0)

      if (!isReportRow && (openedThisWeek || (issue.statusKind !== 'open' && issue.statusKind !== 'pending'))) return

      let group: IssueDetailRow['group'] = 'Open Carryover'
      if (!isReportRow) group = 'Open for Discussion'
      else if (issue.statusKind === 'closed' && openedInReport && closedInReport) group = 'Opened + Closed in Report Week'
      else if (issue.statusKind === 'closed' && issue.closedWeek?.label === currentWeek.label) group = 'Closed This Week'
      else if (issue.statusKind === 'closed' && closedInReport) group = 'Closed in Report Week'
      else if (openedInReport) group = 'Opened in Report Week'
      const row: IssueDetailRow = {
        id: issue.id,
        subtype: issue.subtype,
        status: issue.statusKind === 'pending' ? 'Open' : issue.status,
        title: issue.title,
        contractor: issue.contractor,
        workWeek: issue.createdWeek?.label ?? '',
        createdOn: formatDate(issue.createdOn),
        workWeekClosed: issue.statusKind === 'closed' ? issue.closedWeek?.label ?? '' : 'Open',
        dueDate: formatDate(issue.dueDate),
        group,
      }
      if (group === 'Open for Discussion') discussionRows.push(row)
      else reportRows.push(row)
    })

  const newestFirst = (a: IssueDetailRow, b: IssueDetailRow) =>
    b.id.localeCompare(a.id, undefined, { numeric: true, sensitivity: 'base' })
  return selectIssueDetailExportRows([
    ...reportRows.sort(newestFirst),
    ...discussionRows.sort(newestFirst),
  ])
}

export function selectIssueDetailExportRows(rows: IssueDetailRow[], pageSize = 14): IssueDetailRow[] {
  const reportRows = rows.filter((row) => row.group !== 'Open for Discussion')
  const discussionRows = rows.filter((row) => row.group === 'Open for Discussion')
  const reportPageCount = Math.max(1, Math.ceil(reportRows.length / pageSize))
  const availableDiscussionRows = reportPageCount * pageSize - reportRows.length
  return [...reportRows, ...discussionRows.slice(0, availableDiscussionRows)]
}

function buildElectrical(records: InspectionRecord[], reportWeek: WorkWeek): ElectricalPoint[] {
  const weeks = enumerateWorkWeeks(START_WEEK, reportWeek)
  return weeks.map((week) => {
    const finalRows = records.filter((row) =>
      row.workWeek?.label === week.label && phaseEquals(row.phase, 'Final'),
    )
    return {
      workWeek: week.label,
      finals: finalRows.length,
      issuesFound: finalRows.filter((row) =>
        row.issue && row.issue.trim().toLowerCase() !== 'no issue found',
      ).length,
    }
  })
}

function buildWelding(records: WeldRecord[], reportWeek: WorkWeek): WeldingPoint[] {
  const weeks = enumerateWorkWeeks(START_WEEK, reportWeek)
  return weeks.map((week) => {
    const rows = records.filter((row) => row.workWeek?.label === week.label && row.no.trim())
    const signed = rows.filter((row) => row.signature.trim()).length
    return {
      workWeek: week.label,
      signed,
      total: rows.length,
      signoffRate: rows.length ? (signed / rows.length) * 100 : 0,
      issuesCreated: rows.filter((row) => row.issueCreated.trim()).length,
    }
  })
}

function sheetHealth(bundle: SheetBundle): Record<SheetRole, boolean> {
  return {
    bimIssues: bundle.sheets.bimIssues.rows.length > 0,
    mechanical: bundle.sheets.mechanical.rows.length > 0,
    electrical: bundle.sheets.electrical.rows.length > 0,
    welding: bundle.sheets.welding.rows.length > 0,
  }
}

export function mergeFilters(saved: Partial<ReportFilters>): ReportFilters {
  const reportingMode = saved.reportingMode ?? (saved.oac === false ? 'manual' : 'oac')
  return {
    ...EMPTY_FILTERS,
    ...saved,
    reportingMode,
    oac: reportingMode !== 'manual',
    contractors: reportingMode === 'justine' ? [] : saved.contractors ?? EMPTY_FILTERS.contractors,
  }
}

export function buildReportModel(
  bundle: SheetBundle,
  filtersInput: ReportFilters,
  now = new Date(),
): ReportModel {
  const currentWeek = toIsoWorkWeek(now)
  const reportingMode = filtersInput.reportingMode
  const periodEnd = reportingMode === 'justine'
    ? currentWeek
    : filtersInput.oac ? previousWorkWeek(currentWeek) : currentWeek
  const periodStart = reportingMode === 'justine' ? previousWorkWeek(currentWeek) : periodEnd
  const reportWeek = periodEnd
  const previousReport = previousWorkWeek(periodStart)
  const previousPeriodStart = reportingMode === 'justine'
    ? previousWorkWeek(previousReport)
    : previousReport
  const cutoffDate = reportingMode === 'justine' || !filtersInput.oac ? now : workWeekEnd(reportWeek)
  const periodLabel = periodStart.label === periodEnd.label
    ? periodEnd.label
    : `${periodStart.label} + ${periodEnd.label}`

  const allIssues = bundle.sheets.bimIssues.rows.filter(shouldIncludeIssue).map(normalizeIssue)
  const allMechanical = bundle.sheets.mechanical.rows.map(normalizeInspection)
  const allElectrical = bundle.sheets.electrical.rows.map(normalizeInspection)
  const allWelds = bundle.sheets.welding.rows.map(normalizeWeld)
  const allInspections = [...allMechanical, ...allElectrical]

  const filterOptions = buildFilterOptions(allIssues, allInspections, allWelds)
  const filters = {
    ...filtersInput,
    workWeeks: filtersInput.workWeeks.filter((item) => filterOptions.workWeeks.includes(item)),
    disciplines: filtersInput.disciplines.filter((item) => filterOptions.disciplines.includes(item)),
    contractors: filtersInput.contractors.filter((item) => filterOptions.contractors.includes(item)),
    subtypes: filtersInput.subtypes.filter((item) => filterOptions.subtypes.includes(item)),
    statuses: filtersInput.statuses.filter((item) => filterOptions.statuses.includes(item)),
  }

  const issues = allIssues.filter((issue) => passesIssueFilters(issue, filters, reportWeek))
  const inspections = allInspections.filter((row) => passesInspectionFilters(row, filters, reportWeek))
  const electricalRecords = allElectrical.filter((row) => passesInspectionFilters(row, filters, reportWeek))
  const welds = allWelds.filter((row) => passesWeldFilters(row, filters, reportWeek))

  const totalOpened = countOpenedThrough(issues, reportWeek)
  const previousTotalOpened = countOpenedThrough(issues, previousReport)
  const totalClosed = countClosedThrough(issues, reportWeek)
  const previousTotalClosed = countClosedThrough(issues, previousReport)
  const openedWeek = countOpenedBetween(issues, periodStart, periodEnd)
  const openedPreviousWeek = countOpenedBetween(issues, previousPeriodStart, previousReport)
  const closedWeek = countClosedBetween(issues, periodStart, periodEnd)
  const closedPreviousWeek = countClosedBetween(issues, previousPeriodStart, previousReport)
  const remaining = totalOpened - totalClosed
  const previousRemaining = previousTotalOpened - previousTotalClosed
  const inspectionsWeek = inspections.filter((row) => isWeekInRange(row.workWeek, periodStart, periodEnd) && phaseEquals(row.phase, 'Final')).length
  const inspectionsPrevious = inspections.filter((row) => isWeekInRange(row.workWeek, previousPeriodStart, previousReport) && phaseEquals(row.phase, 'Final')).length
  const sorsWeek = inspections.filter((row) => isWeekInRange(row.workWeek, periodStart, periodEnd) && includesPhase(row.phase, 'SOR')).length
  const sorsPrevious = inspections.filter((row) => isWeekInRange(row.workWeek, previousPeriodStart, previousReport) && includesPhase(row.phase, 'SOR')).length
  const closureRate = totalOpened ? (totalClosed / totalOpened) * 100 : 0
  const previousClosureRate = previousTotalOpened ? (previousTotalClosed / previousTotalOpened) * 100 : 0

  // Per-week history for KPI sparklines (last SPARK_WEEKS reporting weeks).
  const SPARK_WEEKS = 12
  const sparkWeeks = enumerateWorkWeeks(START_WEEK, reportWeek).slice(-SPARK_WEEKS)
  const openedWeekSeries = sparkWeeks.map((week) => countOpenedIn(issues, week))
  const closedWeekSeries = sparkWeeks.map((week) => countClosedIn(issues, week))
  const cumOpenedSeries = sparkWeeks.map((week) => countOpenedThrough(issues, week))
  const cumClosedSeries = sparkWeeks.map((week) => countClosedThrough(issues, week))
  const remainingSeries = cumOpenedSeries.map((opened, i) => opened - cumClosedSeries[i])
  const closureRateSeries = cumOpenedSeries.map((opened, i) => (opened ? (cumClosedSeries[i] / opened) * 100 : 0))
  const inspectionsSeries = sparkWeeks.map((week) => inspections.filter((row) => row.workWeek?.label === week.label && phaseEquals(row.phase, 'Final')).length)
  const sorsSeries = sparkWeeks.map((week) => inspections.filter((row) => row.workWeek?.label === week.label && includesPhase(row.phase, 'SOR')).length)

  const kpis: KpiMetric[] = [
    makeMetric(
      'total-opened',
      'Total Issues Opened',
      totalOpened,
      previousTotalOpened,
      ClipboardList,
      totalOpened > previousTotalOpened ? 'bad' : totalOpened < previousTotalOpened ? 'good' : 'neutral',
      compactNumber,
      cumOpenedSeries,
    ),
    makeMetric('total-closed', 'Total Issues Closed', totalClosed, previousTotalClosed, CheckCircle2, 'good', compactNumber, cumClosedSeries),
    makeMetric('opened-week', 'Issues Opened', openedWeek, openedPreviousWeek, TrendingUp, openedWeek > openedPreviousWeek ? 'warn' : 'neutral', compactNumber, openedWeekSeries),
    makeMetric('closed-week', 'Issues Closed', closedWeek, closedPreviousWeek, TrendingDown, closedWeek >= closedPreviousWeek ? 'good' : 'neutral', compactNumber, closedWeekSeries),
    makeMetric(
      'remaining-open',
      'Issues Remaining Open',
      remaining,
      previousRemaining,
      AlertCircle,
      remaining > previousRemaining ? 'bad' : remaining < previousRemaining ? 'good' : 'neutral',
      compactNumber,
      remainingSeries,
    ),
    makeMetric('inspections', 'Inspections', inspectionsWeek, inspectionsPrevious, ClipboardCheck, 'good', compactNumber, inspectionsSeries),
    makeMetric('sors', 'SORs', sorsWeek, sorsPrevious, Wrench, sorsWeek > sorsPrevious ? 'warn' : 'neutral', compactNumber, sorsSeries),
    makeMetric(
      'closure-rate',
      'Closure Rate',
      closureRate,
      previousClosureRate,
      Gauge,
      closureRate > previousClosureRate ? 'good' : closureRate < previousClosureRate ? 'bad' : 'neutral',
      (value) => percent(value, 1),
      closureRateSeries,
    ),
  ]

  const electrical = buildElectrical(electricalRecords, reportWeek)
  const welding = buildWelding(welds, reportWeek)
  const reportElectricalPoints = electrical.filter((point) => isWeekInRange(parseWorkWeek(point.workWeek), periodStart, periodEnd))
  const previousElectricalPoints = electrical.filter((point) => isWeekInRange(parseWorkWeek(point.workWeek), previousPeriodStart, previousReport))
  const reportWeldPoints = welding.filter((point) => isWeekInRange(parseWorkWeek(point.workWeek), periodStart, periodEnd))
  const previousWeldPoints = welding.filter((point) => isWeekInRange(parseWorkWeek(point.workWeek), previousPeriodStart, previousReport))
  const reportElectrical = {
    finals: reportElectricalPoints.reduce((sum, point) => sum + point.finals, 0),
    issuesFound: reportElectricalPoints.reduce((sum, point) => sum + point.issuesFound, 0),
  }
  const prevElectrical = {
    finals: previousElectricalPoints.reduce((sum, point) => sum + point.finals, 0),
    issuesFound: previousElectricalPoints.reduce((sum, point) => sum + point.issuesFound, 0),
  }
  const reportWeldTotal = reportWeldPoints.reduce((sum, point) => sum + point.total, 0)
  const reportWeldSigned = reportWeldPoints.reduce((sum, point) => sum + point.signed, 0)
  const previousWeldTotal = previousWeldPoints.reduce((sum, point) => sum + point.total, 0)
  const previousWeldSigned = previousWeldPoints.reduce((sum, point) => sum + point.signed, 0)
  const reportWeld = {
    total: reportWeldTotal,
    signed: reportWeldSigned,
    signoffRate: reportWeldTotal ? (reportWeldSigned / reportWeldTotal) * 100 : 0,
  }
  const prevWeld = {
    total: previousWeldTotal,
    signed: previousWeldSigned,
    signoffRate: previousWeldTotal ? (previousWeldSigned / previousWeldTotal) * 100 : 0,
  }
  const weldingWeeksWithData = welding.filter((point) => point.total > 0)
  const overallSignoffRate = weldingWeeksWithData.length
    ? weldingWeeksWithData.reduce((sum, point) => sum + point.signoffRate, 0) / weldingWeeksWithData.length
    : 0

  return {
    generatedAt: now,
    reportingMode,
    currentWeek,
    reportWeek,
    previousReportWeek: previousReport,
    periodStartWeek: periodStart,
    periodEndWeek: periodEnd,
    periodLabel,
    cutoffDate,
    source: bundle.source,
    sheetHealth: sheetHealth(bundle),
    filterOptions,
    activeFilters: filters,
    kpis,
    issueTrend: buildWeeklyTrend(issues, reportWeek),
    openAging: buildOpenAging(issues, now),
    aging: buildAging(issues, now),
    issueTable: buildIssueTable(
      allIssues.filter((issue) => passesIssueFilters(issue, { ...filters, oac: false, workWeeks: [] }, reportWeek)),
      currentWeek,
      periodStart,
      periodEnd,
      filters.oac,
    ),
    electrical,
    welding,
    summary: {
      electricalFinals: reportElectrical?.finals ?? 0,
      electricalIssuesFound: reportElectrical?.issuesFound ?? 0,
      weldsChecked: reportWeld?.total ?? 0,
      weldsSigned: reportWeld?.signed ?? 0,
      overallSignoffRate,
      reportWeekSignoffRate: reportWeld?.signoffRate ?? 0,
      deltas: {
        electricalFinals: (reportElectrical?.finals ?? 0) - (prevElectrical?.finals ?? 0),
        electricalIssuesFound: (reportElectrical?.issuesFound ?? 0) - (prevElectrical?.issuesFound ?? 0),
        weldsChecked: (reportWeld?.total ?? 0) - (prevWeld?.total ?? 0),
        weldsSigned: (reportWeld?.signed ?? 0) - (prevWeld?.signed ?? 0),
        reportWeekSignoffRate: (reportWeld?.signoffRate ?? 0) - (prevWeld?.signoffRate ?? 0),
      },
    },
  }
}
