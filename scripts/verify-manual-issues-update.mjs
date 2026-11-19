import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import { build } from 'vite'

const root = process.cwd()
const temporaryDirectory = await mkdtemp(resolve(tmpdir(), 'qaqc-manual-issues-'))
const entryPath = resolve(temporaryDirectory, 'manual-issues-check.ts')
const outputDirectory = resolve(temporaryDirectory, 'dist')

const entrySource = `
import { File } from 'node:buffer'
import * as XLSX from ${JSON.stringify(resolve(root, 'node_modules/xlsx/xlsx.mjs'))}
import { strFromU8, unzipSync } from ${JSON.stringify(resolve(root, 'node_modules/fflate/esm/index.mjs'))}
import {
  buildUpdatedIssueWorkbook,
  prepareIssueWorkbook,
  reconcileIssueRows,
} from ${JSON.stringify(resolve(root, 'src/services/manualIssuesUpdate.ts'))}

const workbookBytes = (rows, includeNotes = false) => {
  const workbook = XLSX.utils.book_new()
  const worksheet = XLSX.utils.aoa_to_sheet(rows)
  worksheet['!autofilter'] = { ref: 'A1:' + XLSX.utils.encode_col(rows[0].length - 1) + rows.length }
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Issues')
  if (includeNotes) {
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['Keep this sheet'], ['Untouched']]), 'Notes')
  }
  return XLSX.write(workbook, { type: 'array', bookType: 'xlsx', cellStyles: true })
}

const currentBytes = workbookBytes([
  ['ID', 'Contractor', 'Discipline', 'Created By', 'Created On', 'Updated On', 'Status', 'Due Date', 'Reference Notes'],
  ['BIM-100', 'Old Trade', 'BIM', 'Samuel Leach LotusWorks', '2026-01-01', '2026-03-01', 'Pending', '2026-04-01', 'fill both'],
  ['BIM-101', 'Trade A', 'Electrical', 'Samuel Leach LotusWorks', '2026-01-02', '2026-03-02', 'Open', '2026-04-02', 'fill discipline only'],
  ['BIM-102', '', 'Mechanical', 'Samuel Leach LotusWorks', '2026-01-03', '2026-02-01', 'Closed', '', 'dates only'],
  ['BIM-103', 'Trade C', '', 'Samuel Leach LotusWorks', '2026-01-04', '2026-03-04', 'Open', '', 'fill contractor only'],
  ['BIM-104', 'Trade D', 'Piping', '', '2026-01-05', '2026-03-05', 'Open', '', 'first duplicate wins'],
  ['BIM-104', 'Wrong Duplicate', 'Wrong Duplicate', 'Wrong Duplicate', '2025-01-01', '2025-01-02', 'Closed', '2025-04-01', 'ignored duplicate'],
])
const accBytes = workbookBytes([
  ['Issue ID', 'Responsible Contractor', 'Trade', 'Title', 'Status', 'Due date', 'ACC Metadata'],
  ['BIM-100', '', '', 'One hundred', 'Pending', '', 'keep-100'],
  ['BIM-101', 'ACC Trade', '', 'One hundred one', 'Open', '2026-04-20', 'keep-101'],
  ['BIM-102', '', 'ACC Discipline', 'One hundred two', 'Closed', '', 'keep-102'],
  ['BIM-103', '', '', 'One hundred three', 'Open', '', 'keep-103'],
  ['BIM-104', '   ', '   ', 'One hundred four', 'Open', '', 'keep-104'],
  ['BIM-999', '', '', 'Unmatched issue', 'Open', '', 'keep-999'],
  ['', '', '', 'Missing ID row', 'Open', '', 'keep-missing'],
], true)

const currentFile = new File([currentBytes], 'Current_BIM_Issues_Log.xlsx')
const accFile = new File([accBytes], 'ACC_Issues_Export.xlsx')
const current = await prepareIssueWorkbook(currentFile, 'current')
const acc = await prepareIssueWorkbook(accFile, 'acc')
const analysis = reconcileIssueRows(current, acc)

if (analysis.currentRows !== 6 || analysis.accRows !== 7) throw new Error('Workbook row counts changed unexpectedly.')
if (analysis.matchedRows !== 5 || analysis.unmatchedRows !== 1 || analysis.missingIdRows !== 1) {
  throw new Error('ID matching did not preserve unmatched or missing-ID ACC rows.')
}
if (analysis.duplicateCurrentIds !== 1) throw new Error('Duplicate reference IDs were not detected.')
if (analysis.unchangedMatchedRows !== 0) throw new Error('A matched historical row did not receive its provenance metadata.')
if (analysis.filledContractors !== 3 || analysis.filledDisciplines !== 3) {
  throw new Error('Contractor and Discipline fill counts are incorrect.')
}
if (analysis.enrichedCreators !== 5 || analysis.enrichedCreatedDates !== 5 || analysis.enrichedClosedDates !== 1) {
  throw new Error('Historical ownership and driving-date counts are incorrect.')
}
if (analysis.filledDueDates !== 1) throw new Error('Blank ACC Due Date fill count is incorrect.')
if (analysis.changes.map((change) => change.id).join(',') !== 'BIM-100,BIM-101,BIM-102,BIM-103,BIM-104') {
  throw new Error('The wrong ACC rows were selected for enrichment.')
}

const output = buildUpdatedIssueWorkbook(acc, analysis)
const enriched = XLSX.read(output.bytes, { type: 'array', cellDates: true, cellStyles: true })
if (enriched.SheetNames.join(',') !== 'Issues,Notes') throw new Error('The ACC workbook sheet structure was not preserved.')
if (enriched.Sheets.Notes.A1?.v !== 'Keep this sheet' || enriched.Sheets.Notes.A2?.v !== 'Untouched') {
  throw new Error('A secondary ACC worksheet was modified.')
}
if (enriched.Sheets.Issues['!autofilter']?.ref !== 'A1:J8') {
  throw new Error('The ACC worksheet filter range was not preserved.')
}

const rows = XLSX.utils.sheet_to_json(enriched.Sheets.Issues, { defval: '', raw: true })
if (rows.map((row) => row['Issue ID']).join(',') !== 'BIM-100,BIM-101,BIM-102,BIM-103,BIM-104,BIM-999,') {
  throw new Error('ACC row order or row membership changed.')
}
const byId = new Map(rows.filter((row) => row['Issue ID']).map((row) => [row['Issue ID'], row]))
const expectFields = (id, contractor, discipline) => {
  const row = byId.get(id)
  if (row?.['Responsible Contractor'] !== contractor || row?.Trade !== discipline) {
    throw new Error(id + ' did not retain the expected Contractor and Discipline.')
  }
}
expectFields('BIM-100', 'Old Trade', 'BIM')
expectFields('BIM-101', 'ACC Trade', 'Electrical')
expectFields('BIM-102', '', 'ACC Discipline')
expectFields('BIM-103', 'Trade C', '')
expectFields('BIM-104', 'Trade D', 'Piping')
expectFields('BIM-999', '', '')

const isoDate = (value) => value instanceof Date ? value.toISOString().slice(0, 10) : String(value ?? '')
for (const id of ['BIM-100', 'BIM-101', 'BIM-102', 'BIM-103', 'BIM-104']) {
  if (!String(byId.get(id)?.['BIM360_Created By'] ?? '').toLowerCase().includes('lotusworks')) {
    throw new Error(id + ' did not retain historical BIM ownership provenance.')
  }
}
if (isoDate(byId.get('BIM-100')?.['BIM360_Created On']) !== '2026-01-01') {
  throw new Error('Historical Created On was not written as the driving created date.')
}
if (isoDate(byId.get('BIM-102')?.['BIM360_Closed On']) !== '2026-02-01') {
  throw new Error('A closed historical issue did not derive its closure date from Updated On.')
}
if (byId.get('BIM-100')?.['BIM360_Closed On'] !== '' || byId.get('BIM-104')?.['BIM360_Closed On'] !== '') {
  throw new Error('An open historical issue incorrectly received a closed date.')
}
if (isoDate(byId.get('BIM-104')?.['BIM360_Created On']) !== '2026-01-05') {
  throw new Error('The first matching historical ID did not control duplicate-date enrichment.')
}
if (isoDate(byId.get('BIM-100')?.['Due date']) !== '2026-04-01') {
  throw new Error('A blank ACC Due Date was not filled from the matching historical issue.')
}
if (isoDate(byId.get('BIM-101')?.['Due date']) !== '2026-04-20') {
  throw new Error('An existing ACC Due Date was incorrectly overwritten by the historical issue.')
}

rows.forEach((row) => {
  const id = row['Issue ID'] || 'missing'
  const expected = id === 'missing' ? 'keep-missing' : 'keep-' + id.replace('BIM-', '')
  if (row['ACC Metadata'] !== expected) throw new Error(id + ' lost unrelated ACC metadata.')
})
if (!output.fileName.startsWith('ACC_Issues_Export-Enriched-') || !output.fileName.endsWith('.xlsx')) {
  throw new Error('The enriched workbook filename is not based on the ACC export.')
}
if (acc.rows[0]['Responsible Contractor'] !== '' || acc.rows[0].Trade !== '') {
  throw new Error('The in-memory ACC source model was mutated.')
}

const importWorkbook = XLSX.utils.book_new()
const importSheet = XLSX.utils.aoa_to_sheet([])
XLSX.utils.sheet_add_aoa(importSheet, [
  ['Issue ID', 'Contractor', 'Discipline', 'Updated On'],
  [{ t: 's', v: 'BIM-100', f: '"BIM-100"' }, '', '', { t: 'd', v: new Date('2026-03-01T08:30:00Z'), z: 'General' }],
  [],
  ['BIM-101', '', ''],
  [{ t: 'n', v: 42, z: '00000' }, '', ''],
], { origin: 'B3' })
XLSX.utils.book_append_sheet(importWorkbook, importSheet, 'Issues')
const importSource = XLSX.write(importWorkbook, { type: 'array', bookType: 'xlsx', bookSST: true, cellDates: true })
const importAcc = await prepareIssueWorkbook(new File([importSource], 'ACC.xlsx'), 'acc')
const importOutput = buildUpdatedIssueWorkbook(importAcc, reconcileIssueRows(current, importAcc))
const importResult = XLSX.read(importOutput.bytes, { cellNF: true, cellDates: true }).Sheets.Issues
for (const [address, id] of [['B4', 'BIM-100'], ['B6', 'BIM-101'], ['B7', '00042']]) {
  const cell = importResult[address]
  if (cell?.t !== 's' || cell.v !== id || cell.z !== '@' || cell.f || cell.F) {
    throw new Error(address + ' was not exported as a literal text ID.')
  }
}
if (importResult.C4?.v !== 'Old Trade' || importResult.C5 || importResult.C6?.v !== 'Trade A') {
  throw new Error('Enrichment shifted columns or rows around the blank row.')
}
if (importResult.E4?.v?.toISOString() !== '2026-03-01T08:30:00.000Z') {
  throw new Error('An existing date with General formatting lost its date or time.')
}
const zip = unzipSync(new Uint8Array(importOutput.bytes))
const xml = strFromU8(zip['xl/worksheets/sheet1.xml'])
if (!zip['xl/sharedStrings.xml'] || !xml.includes('t="s"') || xml.includes('t="str"') || xml.includes('<f') || xml.includes('t="d"')) {
  throw new Error('The export must use shared strings, literal IDs, and numeric date serials.')
}

const csv = await prepareIssueWorkbook(new File(['ID,Contractor,Discipline\\n00042,Trade A,Electrical'], 'ACC.csv'), 'acc')
const csvOutput = buildUpdatedIssueWorkbook(csv, reconcileIssueRows(current, csv))
if (XLSX.read(csvOutput.bytes).Sheets.Sheet1.A2?.v !== '00042') {
  throw new Error('A CSV ID lost its leading zeros.')
}
`

try {
  await writeFile(entryPath, entrySource)
  await build({
    configFile: false,
    logLevel: 'silent',
    resolve: { alias: { '@': resolve(root, 'src') } },
    ssr: { noExternal: true },
    build: {
      ssr: entryPath,
      target: 'esnext',
      outDir: outputDirectory,
      emptyOutDir: true,
      rollupOptions: { output: { entryFileNames: 'manual-issues-check.mjs', format: 'es' } },
    },
  })
  await import(`${pathToFileURL(resolve(outputDirectory, 'manual-issues-check.mjs')).href}?t=${Date.now()}`)
  console.log('Manual update preserves ACC data while enriching ownership, driving dates, and blank Due Date cells.')
} finally {
  await rm(temporaryDirectory, { recursive: true, force: true })
}
