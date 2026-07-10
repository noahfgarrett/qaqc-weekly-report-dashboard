import * as XLSX from 'xlsx'
import type { ImportedSheetFile, SheetRecord, SheetRole } from '@/types'

export const DEFAULT_ROLE_NAMES: Record<SheetRole, string> = {
  bimIssues: 'BIM Issues Log',
  mechanical: 'Mechanical / Process Inspection Log',
  electrical: 'Electrical Inspection Log',
  welding: 'Welding Signoffs by Work Week',
}

const SUPPORTED_EXTENSIONS = ['.xls', '.xlsx', '.csv']
const KNOWN_HEADERS = new Set([
  'id',
  'status',
  'createdon',
  'updatedon',
  'inspectionphase',
  'workweekobserved',
  'issue',
  'no',
  'weldworkweek',
  'signature',
  'issuecreatedputbimifyes',
])

interface ParsedWorksheet {
  name: string
  rows: Record<string, unknown>[]
  headers: Set<string>
  quality: number
}

function normalized(value: unknown): string {
  return String(value ?? '').toLowerCase().replace(/[^a-z0-9]/g, '')
}

function extension(fileName: string): string {
  const dot = fileName.lastIndexOf('.')
  return dot >= 0 ? fileName.slice(dot).toLowerCase() : ''
}

function headerScore(row: unknown[]): number {
  return row.reduce<number>((score, cell) => score + (KNOWN_HEADERS.has(normalized(cell)) ? 1 : 0), 0)
}

function parseWorksheet(name: string, worksheet: XLSX.WorkSheet): ParsedWorksheet | null {
  const matrix = XLSX.utils.sheet_to_json<unknown[]>(worksheet, {
    header: 1,
    defval: '',
    raw: false,
  })
  if (matrix.length === 0) return null

  let headerRow = 0
  let bestScore = -1
  matrix.slice(0, 20).forEach((row, index) => {
    const score = headerScore(row)
    if (score > bestScore) {
      headerRow = index
      bestScore = score
    }
  })

  const headers = new Set((matrix[headerRow] ?? []).map(normalized).filter(Boolean))
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet, {
    range: headerRow,
    defval: '',
    raw: false,
  }).map((row, index) => ({ ...row, __rowNumber: headerRow + index + 2 }))

  if (rows.length === 0) return null
  return {
    name,
    rows,
    headers,
    quality: bestScore * 100 + Math.min(rows.length, 99),
  }
}

function includesAll(headers: Set<string>, values: string[]): boolean {
  return values.every((header) => headers.has(header))
}

function roleScores(fileName: string, headers: Set<string>): Record<SheetRole, number> {
  const name = normalized(fileName.replace(/\.[^.]+$/, ''))
  const inspectionHeaders = includesAll(headers, ['inspectionphase', 'workweekobserved'])
  const scores: Record<SheetRole, number> = {
    bimIssues: includesAll(headers, ['id', 'status', 'createdon']) ? 12 : 0,
    mechanical: inspectionHeaders ? 7 : 0,
    electrical: inspectionHeaders ? 7 : 0,
    welding: includesAll(headers, ['no', 'weldworkweek', 'signature']) ? 12 : 0,
  }

  if (name.includes('bim') && name.includes('issue')) scores.bimIssues += 14
  if ((name.includes('mechanical') || name.includes('process')) && name.includes('inspection')) scores.mechanical += 14
  if (name.includes('electrical') && name.includes('inspection')) scores.electrical += 14
  if (name.includes('weld')) scores.welding += name.includes('sign') ? 14 : 10
  return scores
}

function identifyRole(fileName: string, headers: Set<string>): SheetRole {
  const scores = roleScores(fileName, headers)
  const ordered = (Object.entries(scores) as Array<[SheetRole, number]>).sort((a, b) => b[1] - a[1])
  const [best, next] = ordered
  if (best[1] < 10) {
    throw new Error(`${fileName}: could not match this export to one of the four required logs.`)
  }
  if (best[1] === next[1]) {
    throw new Error(`${fileName}: inspection exports need "Mechanical", "Process", or "Electrical" in the filename.`)
  }
  return best[0]
}

export async function importSpreadsheet(file: File): Promise<ImportedSheetFile> {
  if (!SUPPORTED_EXTENSIONS.includes(extension(file.name))) {
    throw new Error(`${file.name}: use an .xls, .xlsx, or .csv Smartsheet export.`)
  }

  let workbook: XLSX.WorkBook
  try {
    workbook = XLSX.read(await file.arrayBuffer(), {
      type: 'array',
      cellDates: false,
      dense: true,
    })
  } catch {
    throw new Error(`${file.name}: the spreadsheet could not be read.`)
  }

  const worksheets = workbook.SheetNames
    .map((name) => parseWorksheet(name, workbook.Sheets[name]))
    .filter((sheet): sheet is ParsedWorksheet => sheet !== null)

  if (worksheets.length === 0) {
    throw new Error(`${file.name}: no populated worksheet was found.`)
  }

  const candidates = worksheets.map((worksheet) => {
    const scores = roleScores(file.name, worksheet.headers)
    return {
      worksheet,
      roleScore: Math.max(...Object.values(scores)),
    }
  }).sort((a, b) => b.roleScore - a.roleScore || b.worksheet.quality - a.worksheet.quality)

  const worksheet = candidates[0].worksheet
  const role = identifyRole(file.name, worksheet.headers)
  const sheet: SheetRecord = {
    id: `${file.name}:${worksheet.name}`,
    name: DEFAULT_ROLE_NAMES[role],
    rows: worksheet.rows,
  }

  return {
    role,
    fileName: file.name,
    worksheetName: worksheet.name,
    rowCount: worksheet.rows.length,
    sheet,
  }
}
