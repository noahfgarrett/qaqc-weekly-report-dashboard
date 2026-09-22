import type { SheetBundle, SheetRole } from '@/types'
import { isLotusWorksIssueCreator } from '@/utils/issueOwnership'

export type DataQualityField =
  | 'id'
  | 'status'
  | 'subtype'
  | 'title'
  | 'contractor'
  | 'discipline'
  | 'createdOn'
  | 'closedOn'
  | 'dueDate'
  | 'inspectionPhase'
  | 'workWeekObserved'
  | 'weldNumber'
  | 'weldWorkWeek'

export interface DataQualityFinding {
  id: string
  role: SheetRole
  rowIndex: number
  rowNumber: number
  recordLabel: string
  field: DataQualityField
  fieldLabel: string
  targetColumn: string
  suggestedValue: string
  autoFilled: boolean
}

export interface DataQualityAudit {
  findings: DataQualityFinding[]
  initialValues: Record<string, string>
}

interface FieldDefinition {
  field: DataQualityField
  label: string
  canonicalColumn: string
  aliases: string[]
  suggestedValue?: string
  when?: (row: Record<string, unknown>) => boolean
}

const ROLE_LABELS: Record<SheetRole, string> = {
  bimIssues: 'BIM Issues Log',
  mechanical: 'Mechanical / Process Inspection Log',
  electrical: 'Electrical Inspection Log',
  welding: 'Welding Signoffs',
}

const CONTRACTOR_FIELD: FieldDefinition = {
  field: 'contractor',
  label: 'Contractor',
  canonicalColumn: 'Contractor',
  aliases: ['Contractor', 'General Contractor', 'Responsible Contractor'],
  suggestedValue: 'Bechtel',
}

const ROLE_FIELDS: Record<SheetRole, FieldDefinition[]> = {
  bimIssues: [
    { field: 'id', label: 'ID', canonicalColumn: 'ID', aliases: ['ID', 'Issue ID', 'BIM ID'] },
    { field: 'status', label: 'Status', canonicalColumn: 'Status', aliases: ['Status'] },
    { field: 'subtype', label: 'Type', canonicalColumn: 'Type', aliases: ['Type', 'Subtype', 'Sub Type', 'Issue Subtype'] },
    { field: 'title', label: 'Title', canonicalColumn: 'Title', aliases: ['Title', 'Issue', 'Description'] },
    CONTRACTOR_FIELD,
    { field: 'discipline', label: 'Discipline', canonicalColumn: 'Discipline', aliases: ['Discipline', 'Trade'] },
    {
      field: 'createdOn',
      label: 'Created On',
      canonicalColumn: 'Created On',
      aliases: ['BIM360_Created On', 'BIM360 Created On', 'Created On', 'Created', 'Date Created'],
    },
    {
      field: 'closedOn',
      label: 'Closed On',
      canonicalColumn: 'Closed At',
      aliases: ['BIM360_Closed On', 'BIM360 Closed On', 'Closed At', 'Updated On', 'Updated', 'Closed On', 'Date Closed'],
      when: (row) => readValue(row, ['Status']).trim().toLowerCase() === 'closed',
    },
    { field: 'dueDate', label: 'Due Date', canonicalColumn: 'Due Date', aliases: ['Due Date', 'Due'] },
  ],
  mechanical: [
    {
      field: 'inspectionPhase',
      label: 'Inspection Phase',
      canonicalColumn: 'Inspection Phase',
      aliases: ['Inspection Phase', 'Phase of Inspection', 'Inspection Status', 'Phase'],
    },
    {
      field: 'workWeekObserved',
      label: 'Work Week Observed',
      canonicalColumn: 'Work Week Observed',
      aliases: ['Work Week Observed', 'Observed Work Week', 'Week Observed', 'WW Observed', 'Work Week', 'WW'],
    },
    CONTRACTOR_FIELD,
    { field: 'discipline', label: 'Discipline', canonicalColumn: 'Discipline', aliases: ['Discipline', 'Trade'] },
  ],
  electrical: [
    {
      field: 'inspectionPhase',
      label: 'Inspection Phase',
      canonicalColumn: 'Inspection Phase',
      aliases: ['Inspection Phase', 'Phase of Inspection', 'Inspection Status', 'Phase'],
    },
    {
      field: 'workWeekObserved',
      label: 'Work Week Observed',
      canonicalColumn: 'Work Week Observed',
      aliases: ['Work Week Observed', 'Observed Work Week', 'Week Observed', 'WW Observed', 'Work Week', 'WW'],
    },
    CONTRACTOR_FIELD,
    { field: 'discipline', label: 'Discipline', canonicalColumn: 'Discipline', aliases: ['Discipline', 'Trade'] },
  ],
  welding: [
    { field: 'weldNumber', label: 'Weld Number', canonicalColumn: 'NO', aliases: ['NO', 'No.', 'Weld No', 'Weld Number', 'Weld #', 'Weld ID'] },
    {
      field: 'weldWorkWeek',
      label: 'Weld Work Week',
      canonicalColumn: 'WELD WORK WEEK',
      aliases: ['WELD WORK WEEK', 'Weld Work Week', 'Weld WW', 'Work Week Welded', 'Work Week'],
    },
  ],
}

function normalizedKey(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]/g, '')
}

function matchingColumn(row: Record<string, unknown>, aliases: string[]): string | null {
  const columns = new Map(Object.keys(row).map((key) => [normalizedKey(key), key]))
  for (const alias of aliases) {
    const match = columns.get(normalizedKey(alias))
    if (match) return match
  }
  return null
}

function readValue(row: Record<string, unknown>, aliases: string[]): string {
  for (const alias of aliases) {
    const column = matchingColumn(row, [alias])
    if (!column) continue
    const value = String(row[column] ?? '').trim()
    if (value) return value
  }
  return ''
}

function hasAnyColumn(row: Record<string, unknown>, aliases: string[]): boolean {
  return aliases.some((alias) => matchingColumn(row, [alias]) !== null)
}

function reportableIssue(row: Record<string, unknown>): boolean {
  const legacyCreatedBy = readValue(row, ['BIM360_Created By', 'BIM360 Created By'])
  const hasLegacyData = Boolean(
    legacyCreatedBy
    || readValue(row, ['BIM360_Created On', 'BIM360 Created On'])
    || readValue(row, ['BIM360_Closed On', 'BIM360 Closed On']),
  )
  if (hasLegacyData) return isLotusWorksIssueCreator(legacyCreatedBy)
  const creatorAliases = ['Created By', 'Issue Owner']
  if (!hasAnyColumn(row, creatorAliases)) return true
  return isLotusWorksIssueCreator(readValue(row, creatorAliases))
}

function recordLabel(role: SheetRole, row: Record<string, unknown>, rowNumber: number): string {
  if (role === 'bimIssues') return readValue(row, ['ID', 'Issue ID', 'BIM ID']) || `Row ${rowNumber}`
  if (role === 'welding') return readValue(row, ['NO', 'No.', 'Weld No', 'Weld Number', 'Weld #', 'Weld ID']) || `Row ${rowNumber}`
  return readValue(row, ['ID', 'Inspection ID', 'Title']) || `Row ${rowNumber}`
}

export function auditDataQuality(bundle: SheetBundle): DataQualityAudit {
  const findings: DataQualityFinding[] = []
  const initialValues: Record<string, string> = {}

  ;(Object.keys(ROLE_FIELDS) as SheetRole[]).forEach((role) => {
    bundle.sheets[role].rows.forEach((row, rowIndex) => {
      if (role === 'bimIssues' && !reportableIssue(row)) return
      const rowNumber = Number(row.__rowNumber) || rowIndex + 2
      ROLE_FIELDS[role].forEach((definition) => {
        if (definition.when && !definition.when(row)) return
        if (readValue(row, definition.aliases)) return
        const targetColumn = matchingColumn(row, definition.aliases) ?? definition.canonicalColumn
        const id = `${role}:${rowIndex}:${definition.field}`
        const suggestedValue = definition.suggestedValue ?? ''
        findings.push({
          id,
          role,
          rowIndex,
          rowNumber,
          recordLabel: recordLabel(role, row, rowNumber),
          field: definition.field,
          fieldLabel: definition.label,
          targetColumn,
          suggestedValue,
          autoFilled: Boolean(suggestedValue),
        })
        initialValues[id] = suggestedValue
      })
    })
  })

  return { findings, initialValues }
}

export function applyDataQualityEdits(
  bundle: SheetBundle,
  findings: DataQualityFinding[],
  values: Record<string, string>,
): SheetBundle {
  const rowsByRole = new Map<SheetRole, Record<string, unknown>[]>()
  findings.forEach((finding) => {
    const value = values[finding.id]?.trim()
    if (!value) return
    let rows = rowsByRole.get(finding.role)
    if (!rows) {
      rows = bundle.sheets[finding.role].rows.map((row) => ({ ...row }))
      rowsByRole.set(finding.role, rows)
    }
    rows[finding.rowIndex][finding.targetColumn] = value
  })

  return {
    ...bundle,
    sheets: {
      ...bundle.sheets,
      ...Object.fromEntries(Array.from(rowsByRole, ([role, rows]) => [
        role,
        { ...bundle.sheets[role], rows },
      ])),
    },
  }
}

export function dataQualityRoleLabel(role: SheetRole): string {
  return ROLE_LABELS[role]
}
