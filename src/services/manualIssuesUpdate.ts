import * as XLSX from 'xlsx'
import { parseDate } from '@/utils/workWeeks'

export type ManualIssueWorkbookKind = 'current' | 'acc'

export type ManualIssueField =
  | 'id'
  | 'title'
  | 'status'
  | 'subtype'
  | 'createdOn'
  | 'updatedOn'
  | 'dueDate'
  | 'contractor'
  | 'discipline'

interface IssueFieldDefinition {
  key: ManualIssueField
  label: string
  aliases: string[]
}

const ISSUE_FIELDS: IssueFieldDefinition[] = [
  { key: 'id', label: 'ID', aliases: ['ID', 'Issue ID', 'BIM ID'] },
  { key: 'title', label: 'Title', aliases: ['Title', 'Issue', 'Description'] },
  { key: 'status', label: 'Status', aliases: ['Status'] },
  { key: 'subtype', label: 'Subtype', aliases: ['Subtype', 'Sub Type', 'Issue Subtype'] },
  { key: 'createdOn', label: 'Created on', aliases: ['Created On', 'Created', 'Date Created'] },
  { key: 'updatedOn', label: 'Updated on', aliases: ['Updated On', 'Updated', 'Closed On', 'Date Closed'] },
  { key: 'dueDate', label: 'Due date', aliases: ['Due Date', 'Due'] },
  { key: 'contractor', label: 'Contractor', aliases: ['Contractor', 'Responsible Contractor'] },
  { key: 'discipline', label: 'Discipline', aliases: ['Discipline', 'Trade'] },
]

const CREATOR_ALIASES = ['Created By', 'Issue Owner']
const SUPPORTED_EXTENSIONS = ['.xls', '.xlsx', '.csv']
const SHORT_DATE_FORMAT = 'm/d/yy'
const DATE_FIELDS = new Set<ManualIssueField>(['createdOn', 'updatedOn', 'dueDate'])

export interface PreparedIssueWorkbook {
  kind: ManualIssueWorkbookKind
  fileName: string
  worksheetName: string
  rowCount: number
  headerRow: number
  lastDataRow: number
  data: ArrayBuffer
  rows: Record<string, unknown>[]
  fieldHeaders: Record<ManualIssueField, string>
  fieldColumns: Record<ManualIssueField, number>
  creatorHeader?: string
}

export interface ManualIssueCandidate {
  id: string
  title: string
  status: string
  subtype: string
  createdOn: string
  updatedOn: string
  dueDate: string
  contractor: string
  discipline: string
  sourceRow: number
  targetRow?: number
  changedFields: string[]
  values: Record<ManualIssueField, unknown>
}

export interface ManualIssueAnalysis {
  currentRows: number
  accRows: number
  trackedExistingIds: number
  lotusWorksRows: number
  excludedOtherOwners: number
  unchangedExistingIds: number
  skippedDuplicateIds: number
  skippedMissingIds: number
  newIssues: ManualIssueCandidate[]
  updatedIssues: ManualIssueCandidate[]
}

export interface UpdatedIssueWorkbook {
  fileName: string
  bytes: ArrayBuffer
}

interface WorksheetCandidate {
  worksheetName: string
  headerRow: number
  lastDataRow: number
  matrix: unknown[][]
  rows: Record<string, unknown>[]
  fieldHeaders: Partial<Record<ManualIssueField, string>>
  fieldColumns: Partial<Record<ManualIssueField, number>>
  creatorHeader?: string
  score: number
}

function normalized(value: unknown): string {
  return String(value ?? '').trim().toLowerCase().replace(/[^a-z0-9]/g, '')
}

function idKey(value: unknown): string {
  return String(value ?? '').trim().toLowerCase()
}

function extension(fileName: string): string {
  const dot = fileName.lastIndexOf('.')
  return dot >= 0 ? fileName.slice(dot).toLowerCase() : ''
}

function findHeader(headers: unknown[], aliases: string[]): { name: string; column: number } | null {
  const aliasSet = new Set(aliases.map(normalized))
  const column = headers.findIndex((header) => aliasSet.has(normalized(header)))
  if (column < 0) return null
  return { name: String(headers[column] ?? '').trim(), column }
}

function fieldAliases(field: IssueFieldDefinition, kind: ManualIssueWorkbookKind): string[] {
  if (kind === 'acc' && field.key === 'subtype') return ['Type', ...field.aliases]
  return field.aliases
}

function findLastDataRow(matrix: unknown[][], headerRow: number): number {
  for (let row = matrix.length - 1; row > headerRow; row -= 1) {
    if ((matrix[row] ?? []).some((cell) => String(cell ?? '').trim() !== '')) return row
  }
  return headerRow
}

function inspectWorksheet(
  worksheetName: string,
  worksheet: XLSX.WorkSheet,
  kind: ManualIssueWorkbookKind,
): WorksheetCandidate | null {
  const matrix = XLSX.utils.sheet_to_json<unknown[]>(worksheet, {
    header: 1,
    defval: '',
    raw: false,
  })
  if (matrix.length === 0) return null

  let best: WorksheetCandidate | null = null
  matrix.slice(0, 30).forEach((headers, headerRow) => {
    const fieldHeaders: Partial<Record<ManualIssueField, string>> = {}
    const fieldColumns: Partial<Record<ManualIssueField, number>> = {}
    ISSUE_FIELDS.forEach((field) => {
      const match = findHeader(headers, fieldAliases(field, kind))
      if (!match) return
      fieldHeaders[field.key] = match.name
      fieldColumns[field.key] = match.column
    })
    const creatorAliases = kind === 'current' ? [...CREATOR_ALIASES].reverse() : CREATOR_ALIASES
    const creator = findHeader(headers, creatorAliases)
    const matchedFields = Object.keys(fieldHeaders).length
    const score = matchedFields * 100 + (kind === 'acc' && creator ? 150 : 0)
    if (best && best.score >= score) return

    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet, {
      range: headerRow,
      defval: '',
      raw: true,
      blankrows: false,
    }).map((row, index) => ({ ...row, __rowNumber: headerRow + index + 2 }))

    best = {
      worksheetName,
      headerRow,
      lastDataRow: findLastDataRow(matrix, headerRow),
      matrix,
      rows,
      fieldHeaders,
      fieldColumns,
      creatorHeader: creator?.name,
      score,
    }
  })
  return best
}

function requireFieldMap(
  candidate: WorksheetCandidate,
  fileName: string,
): { headers: Record<ManualIssueField, string>; columns: Record<ManualIssueField, number> } {
  const missing = ISSUE_FIELDS.filter((field) => candidate.fieldHeaders[field.key] === undefined)
  if (missing.length > 0) {
    throw new Error(`${fileName}: missing ${missing.map((field) => field.label).join(', ')}.`)
  }
  return {
    headers: candidate.fieldHeaders as Record<ManualIssueField, string>,
    columns: candidate.fieldColumns as Record<ManualIssueField, number>,
  }
}

export async function prepareIssueWorkbook(
  file: File,
  kind: ManualIssueWorkbookKind,
): Promise<PreparedIssueWorkbook> {
  if (!SUPPORTED_EXTENSIONS.includes(extension(file.name))) {
    throw new Error(`${file.name}: use an .xls, .xlsx, or .csv issue export.`)
  }

  const data = await file.arrayBuffer()
  let workbook: XLSX.WorkBook
  try {
    workbook = XLSX.read(data, {
      type: 'array',
      cellDates: true,
      cellStyles: true,
    })
  } catch {
    throw new Error(`${file.name}: the spreadsheet could not be read.`)
  }

  const candidates = workbook.SheetNames
    .map((worksheetName) => inspectWorksheet(worksheetName, workbook.Sheets[worksheetName], kind))
    .filter((candidate): candidate is WorksheetCandidate => candidate !== null)
    .sort((a, b) => b.score - a.score || b.rows.length - a.rows.length)

  const candidate = candidates[0]
  if (!candidate) throw new Error(`${file.name}: no populated worksheet was found.`)
  const fields = requireFieldMap(candidate, file.name)
  if (kind === 'acc' && !candidate.creatorHeader) {
    throw new Error(`${file.name}: missing Created By. Issue Owner is also accepted.`)
  }

  return {
    kind,
    fileName: file.name,
    worksheetName: candidate.worksheetName,
    rowCount: candidate.rows.length,
    headerRow: candidate.headerRow,
    lastDataRow: candidate.lastDataRow,
    data,
    rows: candidate.rows,
    fieldHeaders: fields.headers,
    fieldColumns: fields.columns,
    creatorHeader: candidate.creatorHeader,
  }
}

function rowValue(
  workbook: PreparedIssueWorkbook,
  row: Record<string, unknown>,
  field: ManualIssueField,
): unknown {
  return row[workbook.fieldHeaders[field]] ?? ''
}

function displayValue(value: unknown): string {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: '2-digit' })
  }
  return String(value ?? '').trim()
}

function candidateFromValues(
  values: Record<ManualIssueField, unknown>,
  sourceRow: number,
  changedFields: string[],
  targetRow?: number,
): ManualIssueCandidate {
  return {
    id: displayValue(values.id),
    title: displayValue(values.title),
    status: displayValue(values.status),
    subtype: displayValue(values.subtype),
    createdOn: displayValue(values.createdOn),
    updatedOn: displayValue(values.updatedOn),
    dueDate: displayValue(values.dueDate),
    contractor: displayValue(values.contractor),
    discipline: displayValue(values.discipline),
    sourceRow,
    targetRow,
    changedFields,
    values,
  }
}

function valuesFromRow(workbook: PreparedIssueWorkbook, row: Record<string, unknown>): Record<ManualIssueField, unknown> {
  return Object.fromEntries(
    ISSUE_FIELDS.map((field) => [field.key, rowValue(workbook, row, field.key)]),
  ) as Record<ManualIssueField, unknown>
}

function isBlank(value: unknown): boolean {
  return value === null || value === undefined || String(value).trim() === ''
}

function comparable(value: unknown): string | number | boolean {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.getTime()
  if (typeof value === 'number' || typeof value === 'boolean') return value
  return String(value ?? '').trim()
}

export function reconcileIssueRows(
  current: PreparedIssueWorkbook,
  acc: PreparedIssueWorkbook,
): ManualIssueAnalysis {
  const currentById = new Map<string, Record<string, unknown>>()
  current.rows.forEach((row) => {
    const key = idKey(rowValue(current, row, 'id'))
    if (!key || currentById.has(key)) return
    currentById.set(key, row)
  })

  const processedAccIds = new Set<string>()
  const newIssues: ManualIssueCandidate[] = []
  const updatedIssues: ManualIssueCandidate[] = []
  let lotusWorksRows = 0
  let excludedOtherOwners = 0
  let unchangedExistingIds = 0
  let skippedDuplicateIds = 0
  let skippedMissingIds = 0

  acc.rows.forEach((row) => {
    const incomingValues = valuesFromRow(acc, row)
    const key = idKey(incomingValues.id)
    if (!key) {
      skippedMissingIds += 1
      return
    }
    if (processedAccIds.has(key)) {
      skippedDuplicateIds += 1
      return
    }
    processedAccIds.add(key)

    const currentRow = currentById.get(key)
    const accOwnerIsLotusWorks = normalized(row[acc.creatorHeader as string]).includes('lotusworks')
    if (!currentRow && !accOwnerIsLotusWorks) {
      excludedOtherOwners += 1
      return
    }
    lotusWorksRows += 1

    if (!currentRow) {
      newIssues.push(candidateFromValues(
        incomingValues,
        Number(row.__rowNumber ?? 0),
        ISSUE_FIELDS.map((field) => field.label),
      ))
      return
    }

    const currentValues = valuesFromRow(current, currentRow)
    const mergedValues = { ...currentValues }
    const changedFields: string[] = []
    ISSUE_FIELDS.forEach((field) => {
      const incoming = incomingValues[field.key]
      if (field.key === 'id' || field.key === 'createdOn' || isBlank(incoming)) return
      if (comparable(currentValues[field.key]) !== comparable(incoming)) {
        mergedValues[field.key] = incoming
        changedFields.push(field.label)
      }
    })
    if (changedFields.length === 0) {
      unchangedExistingIds += 1
      return
    }
    const sourceRow = Number(row.__rowNumber ?? 0)
    const currentExcelRow = Number(currentRow.__rowNumber ?? 0)
    updatedIssues.push(candidateFromValues(
      mergedValues,
      sourceRow,
      changedFields,
      currentExcelRow > 0 ? currentExcelRow - 1 : undefined,
    ))
  })

  return {
    currentRows: current.rows.length,
    accRows: acc.rows.length,
    trackedExistingIds: currentById.size,
    lotusWorksRows,
    excludedOtherOwners,
    unchangedExistingIds,
    skippedDuplicateIds,
    skippedMissingIds,
    newIssues,
    updatedIssues,
  }
}

function cloneStyle<T>(value: T): T {
  if (!value || typeof value !== 'object') return value
  return JSON.parse(JSON.stringify(value)) as T
}

function excelDate(value: unknown): Date | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value
  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value)
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d, 12)
  }
  return parseDate(value)
}

function makeCell(value: unknown, template?: XLSX.CellObject, dateField = false): XLSX.CellObject {
  const dateValue = dateField && !isBlank(value) ? excelDate(value) : null
  const cell: XLSX.CellObject = dateValue
    ? { t: 'd', v: dateValue }
    : value instanceof Date && !Number.isNaN(value.getTime())
    ? { t: 'd', v: value }
    : typeof value === 'number'
      ? { t: 'n', v: value }
      : typeof value === 'boolean'
        ? { t: 'b', v: value }
        : { t: 's', v: String(value ?? '') }
  if (template?.s) cell.s = cloneStyle(template.s)
  if (template?.z) cell.z = template.z
  if (dateField) cell.z = SHORT_DATE_FORMAT
  return cell
}

function sortRowsByDescendingId(
  worksheet: XLSX.WorkSheet,
  current: PreparedIssueWorkbook,
  lastDataRow: number,
): void {
  const range = XLSX.utils.decode_range(worksheet['!ref'] ?? 'A1')
  const firstDataRow = current.headerRow + 1
  const idColumn = current.fieldColumns.id
  const rows = Array.from({ length: Math.max(0, lastDataRow - firstDataRow + 1) }, (_, index) => {
    const row = firstDataRow + index
    const cells = new Map<number, XLSX.CellObject>()
    for (let column = range.s.c; column <= range.e.c; column += 1) {
      const cell = worksheet[XLSX.utils.encode_cell({ r: row, c: column })]
      if (cell) cells.set(column, cell)
    }
    const id = String(worksheet[XLSX.utils.encode_cell({ r: row, c: idColumn })]?.v ?? '').trim()
    return { id, originalRow: row, cells }
  })

  rows.sort((a, b) => {
    if (!a.id && !b.id) return a.originalRow - b.originalRow
    if (!a.id) return 1
    if (!b.id) return -1
    return b.id.localeCompare(a.id, undefined, { numeric: true, sensitivity: 'base' })
  })

  for (let row = firstDataRow; row <= lastDataRow; row += 1) {
    for (let column = range.s.c; column <= range.e.c; column += 1) {
      delete worksheet[XLSX.utils.encode_cell({ r: row, c: column })]
    }
  }
  rows.forEach((rowData, index) => {
    const targetRow = firstDataRow + index
    rowData.cells.forEach((cell, column) => {
      worksheet[XLSX.utils.encode_cell({ r: targetRow, c: column })] = cell
    })
  })
}

function formatDateColumns(
  worksheet: XLSX.WorkSheet,
  current: PreparedIssueWorkbook,
  lastDataRow: number,
): void {
  for (let row = current.headerRow + 1; row <= lastDataRow; row += 1) {
    DATE_FIELDS.forEach((field) => {
      const address = XLSX.utils.encode_cell({ r: row, c: current.fieldColumns[field] })
      const cell = worksheet[address]
      if (!cell || isBlank(cell.v)) return
      if (cell.f) {
        cell.z = SHORT_DATE_FORMAT
        delete cell.w
        return
      }
      worksheet[address] = makeCell(cell.v, cell, true)
    })
  }
}

function updatedFileName(fileName: string): string {
  const stem = fileName.replace(/\.[^.]+$/, '') || 'BIM-Issues-Log'
  const date = new Date().toISOString().slice(0, 10)
  return `${stem}-Updated-${date}.xlsx`
}

export function buildUpdatedIssueWorkbook(
  current: PreparedIssueWorkbook,
  analysis: ManualIssueAnalysis,
): UpdatedIssueWorkbook {
  const workbook = XLSX.read(current.data, {
    type: 'array',
    cellDates: true,
    cellStyles: true,
  })
  const worksheet = workbook.Sheets[current.worksheetName]
  if (!worksheet) throw new Error('The BIM Issues Log worksheet is no longer available.')

  analysis.updatedIssues.forEach((issue) => {
    if (issue.targetRow === undefined) return
    ISSUE_FIELDS.forEach((field) => {
      const column = current.fieldColumns[field.key]
      const targetAddress = XLSX.utils.encode_cell({ r: issue.targetRow as number, c: column })
      worksheet[targetAddress] = makeCell(issue.values[field.key], worksheet[targetAddress], DATE_FIELDS.has(field.key))
    })
  })

  analysis.newIssues.forEach((issue, index) => {
    const targetRow = current.lastDataRow + 1 + index
    ISSUE_FIELDS.forEach((field) => {
      const column = current.fieldColumns[field.key]
      const targetAddress = XLSX.utils.encode_cell({ r: targetRow, c: column })
      const templateAddress = XLSX.utils.encode_cell({ r: current.lastDataRow, c: column })
      worksheet[targetAddress] = makeCell(issue.values[field.key], worksheet[templateAddress], DATE_FIELDS.has(field.key))
    })
  })

  const originalRange = XLSX.utils.decode_range(worksheet['!ref'] ?? 'A1')
  const lastDataRow = current.lastDataRow + analysis.newIssues.length
  originalRange.e.r = Math.max(originalRange.e.r, lastDataRow)
  worksheet['!ref'] = XLSX.utils.encode_range(originalRange)
  const autoFilter = worksheet['!autofilter'] as { ref?: string } | undefined
  if (autoFilter?.ref) {
    const filterRange = XLSX.utils.decode_range(autoFilter.ref)
    filterRange.e.r = Math.max(filterRange.e.r, lastDataRow)
    autoFilter.ref = XLSX.utils.encode_range(filterRange)
  }

  sortRowsByDescendingId(worksheet, current, lastDataRow)
  formatDateColumns(worksheet, current, lastDataRow)

  const output = XLSX.write(workbook, {
    type: 'array',
    bookType: 'xlsx',
    cellDates: true,
    cellStyles: true,
    compression: true,
  }) as ArrayBuffer
  return {
    fileName: updatedFileName(current.fileName),
    bytes: output,
  }
}

export function downloadUpdatedIssueWorkbook(output: UpdatedIssueWorkbook): void {
  const blob = new Blob([output.bytes], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  })
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = output.fileName
  document.body.appendChild(anchor)
  anchor.click()
  anchor.remove()
  window.setTimeout(() => URL.revokeObjectURL(url), 1_000)
}
