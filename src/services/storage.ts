import type { ReportFilters, SheetRole } from '@/types'

const NAMESPACE = 'qaqc-report-dashboard'
const TOKEN_KEY = `${NAMESPACE}:smartsheet-token`
const TOKEN_META_KEY = `${NAMESPACE}:token-meta`
const SHEET_MAP_KEY = `${NAMESPACE}:sheet-map`
const FILTERS_KEY = `${NAMESPACE}:filters`

interface TokenMeta {
  rememberedAt: string
  masked: string
}

export function maskToken(token: string): string {
  if (token.length <= 10) return 'saved token'
  return `${token.slice(0, 4)}...${token.slice(-4)}`
}

export function saveSmartsheetToken(token: string): void {
  localStorage.setItem(TOKEN_KEY, token)
  const meta: TokenMeta = {
    rememberedAt: new Date().toISOString(),
    masked: maskToken(token),
  }
  localStorage.setItem(TOKEN_META_KEY, JSON.stringify(meta))
}

export function loadSmartsheetToken(): string | null {
  return localStorage.getItem(TOKEN_KEY)
}

export function loadTokenMeta(): TokenMeta | null {
  const raw = localStorage.getItem(TOKEN_META_KEY)
  if (!raw) return null
  try {
    return JSON.parse(raw) as TokenMeta
  } catch {
    return null
  }
}

export function clearSmartsheetToken(): void {
  localStorage.removeItem(TOKEN_KEY)
  localStorage.removeItem(TOKEN_META_KEY)
}

export function saveSheetMapping(mapping: Partial<Record<SheetRole, number>>): void {
  localStorage.setItem(SHEET_MAP_KEY, JSON.stringify(mapping))
}

export function loadSheetMapping(): Partial<Record<SheetRole, number>> {
  const raw = localStorage.getItem(SHEET_MAP_KEY)
  if (!raw) return {}
  try {
    return JSON.parse(raw) as Partial<Record<SheetRole, number>>
  } catch {
    return {}
  }
}

export function saveFilters(filters: ReportFilters): void {
  localStorage.setItem(FILTERS_KEY, JSON.stringify(filters))
}

export function loadFilters(): Partial<ReportFilters> {
  const raw = localStorage.getItem(FILTERS_KEY)
  if (!raw) return {}
  try {
    return JSON.parse(raw) as Partial<ReportFilters>
  } catch {
    return {}
  }
}
