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
import {
  buildUpdatedIssueWorkbook,
  prepareIssueWorkbook,
  reconcileIssueRows,
} from ${JSON.stringify(resolve(root, 'src/services/manualIssuesUpdate.ts'))}

const workbookBytes = (rows, includeNotes = false) => {
  const workbook = XLSX.utils.book_new()
  const worksheet = XLSX.utils.aoa_to_sheet(rows)
  worksheet['!autofilter'] = { ref: 'A1:F' + rows.length }
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Issues')
  if (includeNotes) {
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['Keep this sheet'], ['Untouched']]), 'Notes')
  }
  return XLSX.write(workbook, { type: 'array', bookType: 'xlsx', cellStyles: true })
}

const currentBytes = workbookBytes([
  ['ID', 'Contractor', 'Discipline', 'Reference Notes'],
  ['BIM-100', 'Old Trade', 'BIM', 'fill both'],
  ['BIM-101', 'Trade A', 'Electrical', 'fill discipline only'],
  ['BIM-102', '', 'Mechanical', 'nothing needed'],
  ['BIM-103', 'Trade C', '', 'fill contractor only'],
  ['BIM-104', 'Trade D', 'Piping', 'first duplicate wins'],
  ['BIM-104', 'Wrong Duplicate', 'Wrong Duplicate', 'ignored duplicate'],
])
const accBytes = workbookBytes([
  ['Issue ID', 'Responsible Contractor', 'Trade', 'Title', 'Status', 'ACC Metadata'],
  ['BIM-100', '', '', 'One hundred', 'Pending', 'keep-100'],
  ['BIM-101', 'ACC Trade', '', 'One hundred one', 'Open', 'keep-101'],
  ['BIM-102', '', 'ACC Discipline', 'One hundred two', 'Closed', 'keep-102'],
  ['BIM-103', '', '', 'One hundred three', 'Open', 'keep-103'],
  ['BIM-104', '   ', '   ', 'One hundred four', 'Open', 'keep-104'],
  ['BIM-999', '', '', 'Unmatched issue', 'Open', 'keep-999'],
  ['', '', '', 'Missing ID row', 'Open', 'keep-missing'],
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
if (analysis.unchangedMatchedRows !== 1) throw new Error('The already-complete matched row was not left unchanged.')
if (analysis.filledContractors !== 3 || analysis.filledDisciplines !== 3) {
  throw new Error('Contractor and Discipline fill counts are incorrect.')
}
if (analysis.changes.map((change) => change.id).join(',') !== 'BIM-100,BIM-101,BIM-103,BIM-104') {
  throw new Error('The wrong ACC rows were selected for enrichment.')
}

const output = buildUpdatedIssueWorkbook(acc, analysis)
const enriched = XLSX.read(output.bytes, { type: 'array', cellStyles: true })
if (enriched.SheetNames.join(',') !== 'Issues,Notes') throw new Error('The ACC workbook sheet structure was not preserved.')
if (enriched.Sheets.Notes.A1?.v !== 'Keep this sheet' || enriched.Sheets.Notes.A2?.v !== 'Untouched') {
  throw new Error('A secondary ACC worksheet was modified.')
}
if (enriched.Sheets.Issues['!autofilter']?.ref !== 'A1:F8') {
  throw new Error('The ACC worksheet filter range was not preserved.')
}

const rows = XLSX.utils.sheet_to_json(enriched.Sheets.Issues, { defval: '', raw: false })
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
  console.log('Manual update fills only blank ACC Contractor and Discipline cells while preserving the ACC workbook.')
} finally {
  await rm(temporaryDirectory, { recursive: true, force: true })
}
