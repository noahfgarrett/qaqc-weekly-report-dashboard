import * as XLSX from 'xlsx'

export type ManualIssueWorkbookKind = 'current' | 'acc'
export type ManualIssueField = 'id' | 'contractor' | 'discipline'
export type EnrichedIssueField = 'contractor' | 'discipline'

interface IssueFieldDefinition {
  key: ManualIssueField
  label: string
  aliases: string[]
}

const ISSUE_FIELDS: IssueFieldDefinition[] = [
  { key: 'id', label: 'ID', aliases: ['ID', 'Issue ID', 'BIM ID'] },
  { key: 'contractor', label: 'Contractor', aliases: ['Contractor', 'Responsible Contractor'] },
  { key: 'discipline', label: 'Discipline', aliases: ['Discipline', 'Trade'] },
]

const SUPPORTED_EXTENSIONS = ['.xls', '.xlsx', '.csv']

export interface PreparedIssueWorkbook {
  kind: ManualIssueWorkbookKind
  fileName: string
  worksheetName: string
  rowCount: number
  headerRow: number
  data: ArrayBuffer
  rows: Record<string, unknown>[]
  fieldHeaders: Record<ManualIssueField, string>
  fieldColumns: Record<ManualIssueField, number>
}

export interface ManualIssueEnrichment {
  id: string
  targetRow: number
  contractor: string
  discipline: string
  filledFields: EnrichedIssueField[]
}

export interface ManualIssueAnalysis {
  currentRows: number
  accRows: number
  matchedRows: number
  unmatchedRows: number
  unchangedMatchedRows: number
  missingIdRows: number
  duplicateCurrentIds: number
  filledContractors: number
  filledDisciplines: number
  changes: ManualIssueEnrichment[]
}

export interface UpdatedIssueWorkbook {
  fileName: string
  bytes: ArrayBuffer
}

interface WorksheetCandidate {
  worksheetName: string
  headerRow: number
  rows: Record<string, unknown>[]
  fieldHeaders: Partial<Record<ManualIssueField, string>>
  fieldColumns: Partial<Record<ManualIssueField, number>>
  score: number
}

function normalized(value: unknown): string {
  return String(value ?? '').trim().toLowerCase().replace(/[^a-z0-9]/g, '')
}

function idKey(value: unknown): string {
  return String(value ?? '').trim().toLowerCase()
}

function isBlank(value: unknown): boolean {
  return value === null || value === undefined || String(value).trim() === ''
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

function inspectWorksheet(
  worksheetName: string,
  worksheet: XLSX.WorkSheet,
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
      const match = findHeader(headers, field.aliases)
      if (!match) return
      fieldHeaders[field.key] = match.name
      fieldColumns[field.key] = match.column
    })
    const score = Object.keys(fieldHeaders).length * 100
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
      rows,
      fieldHeaders,
      fieldColumns,
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
    .map((worksheetName) => inspectWorksheet(worksheetName, workbook.Sheets[worksheetName]))
    .filter((candidate): candidate is WorksheetCandidate => candidate !== null)
    .sort((a, b) => b.score - a.score || b.rows.length - a.rows.length)

  const candidate = candidates[0]
  if (!candidate) throw new Error(`${file.name}: no populated worksheet was found.`)
  const fields = requireFieldMap(candidate, file.name)

  return {
    kind,
    fileName: file.name,
    worksheetName: candidate.worksheetName,
    rowCount: candidate.rows.length,
    headerRow: candidate.headerRow,
    data,
    rows: candidate.rows,
    fieldHeaders: fields.headers,
    fieldColumns: fields.columns,
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
  return String(value ?? '').trim()
}

export function reconcileIssueRows(
  current: PreparedIssueWorkbook,
  acc: PreparedIssueWorkbook,
): ManualIssueAnalysis {
  const currentById = new Map<string, Record<string, unknown>>()
  let duplicateCurrentIds = 0
  current.rows.forEach((row) => {
    const key = idKey(rowValue(current, row, 'id'))
    if (!key) return
    if (currentById.has(key)) {
      duplicateCurrentIds += 1
      return
    }
    currentById.set(key, row)
  })

  const changes: ManualIssueEnrichment[] = []
  let matchedRows = 0
  let unmatchedRows = 0
  let unchangedMatchedRows = 0
  let missingIdRows = 0
  let filledContractors = 0
  let filledDisciplines = 0

  acc.rows.forEach((row) => {
    const id = displayValue(rowValue(acc, row, 'id'))
    const key = idKey(id)
    if (!key) {
      missingIdRows += 1
      return
    }

    const currentRow = currentById.get(key)
    if (!currentRow) {
      unmatchedRows += 1
      return
    }
    matchedRows += 1

    const currentContractor = displayValue(rowValue(current, currentRow, 'contractor'))
    const currentDiscipline = displayValue(rowValue(current, currentRow, 'discipline'))
    const accContractor = displayValue(rowValue(acc, row, 'contractor'))
    const accDiscipline = displayValue(rowValue(acc, row, 'discipline'))
    const filledFields: EnrichedIssueField[] = []

    if (isBlank(accContractor) && !isBlank(currentContractor)) {
      filledFields.push('contractor')
      filledContractors += 1
    }
    if (isBlank(accDiscipline) && !isBlank(currentDiscipline)) {
      filledFields.push('discipline')
      filledDisciplines += 1
    }

    if (filledFields.length === 0) {
      unchangedMatchedRows += 1
      return
    }

    changes.push({
      id,
      targetRow: Number(row.__rowNumber ?? 0) - 1,
      contractor: filledFields.includes('contractor') ? currentContractor : accContractor,
      discipline: filledFields.includes('discipline') ? currentDiscipline : accDiscipline,
      filledFields,
    })
  })

  return {
    currentRows: current.rows.length,
    accRows: acc.rows.length,
    matchedRows,
    unmatchedRows,
    unchangedMatchedRows,
    missingIdRows,
    duplicateCurrentIds,
    filledContractors,
    filledDisciplines,
    changes,
  }
}

function cloneStyle<T>(value: T): T {
  if (!value || typeof value !== 'object') return value
  return JSON.parse(JSON.stringify(value)) as T
}

function makeCell(value: string, template?: XLSX.CellObject): XLSX.CellObject {
  const cell: XLSX.CellObject = { t: 's', v: value }
  if (template?.s) cell.s = cloneStyle(template.s)
  if (template?.z) cell.z = template.z
  return cell
}

function enrichedFileName(fileName: string): string {
  const stem = fileName.replace(/\.[^.]+$/, '') || 'ACC-Issues-Export'
  const date = new Date().toISOString().slice(0, 10)
  return `${stem}-Enriched-${date}.xlsx`
}

export function buildUpdatedIssueWorkbook(
  acc: PreparedIssueWorkbook,
  analysis: ManualIssueAnalysis,
): UpdatedIssueWorkbook {
  const workbook = XLSX.read(acc.data, {
    type: 'array',
    cellDates: true,
    cellStyles: true,
  })
  const worksheet = workbook.Sheets[acc.worksheetName]
  if (!worksheet) throw new Error('The ACC Issues Export worksheet is no longer available.')

  analysis.changes.forEach((change) => {
    change.filledFields.forEach((field) => {
      const column = acc.fieldColumns[field]
      const address = XLSX.utils.encode_cell({ r: change.targetRow, c: column })
      worksheet[address] = makeCell(change[field], worksheet[address])
    })
  })

  const output = XLSX.write(workbook, {
    type: 'array',
    bookType: 'xlsx',
    cellDates: true,
    cellStyles: true,
    compression: true,
  }) as ArrayBuffer

  return {
    fileName: enrichedFileName(acc.fileName),
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
  anchor.style.display = 'none'
  document.body.appendChild(anchor)
  anchor.click()
  anchor.remove()
  window.setTimeout(() => URL.revokeObjectURL(url), 1_000)
}
