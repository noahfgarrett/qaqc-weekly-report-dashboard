import type { SheetBundle, SheetRecord, SheetRole, SmartSheetSummary } from '@/types'

const API_BASE = 'https://api.smartsheet.com/2.0'

interface SmartsheetSheetListResponse {
  data?: SmartSheetSummary[]
}

interface SmartsheetColumn {
  id: number
  title: string
}

interface SmartsheetCell {
  columnId: number
  value?: unknown
  displayValue?: string
}

interface SmartsheetRow {
  id: number
  rowNumber?: number
  cells?: SmartsheetCell[]
}

interface SmartsheetSheetResponse {
  id: number
  name: string
  columns?: SmartsheetColumn[]
  rows?: SmartsheetRow[]
}

export const DEFAULT_ROLE_NAMES: Record<SheetRole, string> = {
  bimIssues: 'BIM Issues Log',
  mechanical: 'Mechanical / Process Inspection Log',
  electrical: 'Electrical Inspection Log',
  welding: 'Welding Signoffs by Work Week',
}

function headers(token: string): HeadersInit {
  return {
    Authorization: `Bearer ${token}`,
    Accept: 'application/json',
  }
}

async function fetchJson<T>(url: string, token: string): Promise<T> {
  const res = await fetch(url, { headers: headers(token) })
  if (!res.ok) {
    const message = await res.text().catch(() => '')
    throw new Error(`Smartsheet request failed (${res.status}). ${message}`)
  }
  return res.json() as Promise<T>
}

export async function listSmartsheets(token: string): Promise<SmartSheetSummary[]> {
  const data = await fetchJson<SmartsheetSheetListResponse>(
    `${API_BASE}/sheets?includeAll=true`,
    token,
  )
  return [...(data.data ?? [])].sort((a, b) => a.name.localeCompare(b.name))
}

export function autoMapSheets(sheets: SmartSheetSummary[]): Partial<Record<SheetRole, number>> {
  const mapping: Partial<Record<SheetRole, number>> = {}
  const normalized = sheets.map((sheet) => ({
    ...sheet,
    normalized: sheet.name.toLowerCase().replace(/\s+/g, ' ').trim(),
  }))

  ;(Object.keys(DEFAULT_ROLE_NAMES) as SheetRole[]).forEach((role) => {
    const expected = DEFAULT_ROLE_NAMES[role].toLowerCase().replace(/\s+/g, ' ').trim()
    const exact = normalized.find((sheet) => sheet.normalized === expected)
    const contains = normalized.find((sheet) => sheet.normalized.includes(expected))
    const loose = normalized.find((sheet) => {
      if (role === 'mechanical') return sheet.normalized.includes('mechanical') && sheet.normalized.includes('inspection')
      if (role === 'electrical') return sheet.normalized.includes('electrical') && sheet.normalized.includes('inspection')
      if (role === 'bimIssues') return sheet.normalized.includes('bim') && sheet.normalized.includes('issue')
      return sheet.normalized.includes('weld') && sheet.normalized.includes('sign')
    })
    const match = exact ?? contains ?? loose
    if (match) mapping[role] = match.id
  })

  return mapping
}

export async function getSheet(token: string, sheetId: number): Promise<SheetRecord> {
  const sheet = await fetchJson<SmartsheetSheetResponse>(
    `${API_BASE}/sheets/${sheetId}?includeAll=true`,
    token,
  )
  const columnTitles = new Map((sheet.columns ?? []).map((column) => [column.id, column.title]))
  const rows = (sheet.rows ?? []).map((row) => {
    const record: Record<string, unknown> = {
      __rowId: row.id,
      __rowNumber: row.rowNumber,
    }
    row.cells?.forEach((cell) => {
      const title = columnTitles.get(cell.columnId)
      if (title) record[title] = cell.displayValue ?? cell.value ?? ''
    })
    return record
  })
  return {
    id: sheet.id,
    name: sheet.name,
    rows,
  }
}

export async function loadMappedSheets(
  token: string,
  mapping: Partial<Record<SheetRole, number>>,
): Promise<Partial<SheetBundle['sheets']>> {
  const entries = await Promise.all(
    (Object.entries(mapping) as Array<[SheetRole, number]>).map(async ([role, sheetId]) => {
      const sheet = await getSheet(token, sheetId)
      return [role, sheet] as const
    }),
  )
  return Object.fromEntries(entries) as Partial<SheetBundle['sheets']>
}
