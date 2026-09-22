// ACC replaces the company name with "Removed" when these LotusWorks accounts are deactivated.
export const REMOVED_LOTUSWORKS_CREATORS = [
  'Samuel Leach',
  'Aaron Harwood',
] as const

function normalizedCreator(value: unknown): string {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
}

export function isLotusWorksIssueCreator(value: unknown): boolean {
  const normalized = normalizedCreator(value)
  if (normalized.includes('lotusworks')) return true
  if (!normalized.includes('removed')) return false
  return REMOVED_LOTUSWORKS_CREATORS.some((name) => normalized.includes(normalizedCreator(name)))
}
