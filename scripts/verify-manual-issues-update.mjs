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

const issueHeaders = [
  'ID', 'Title', 'Status', 'Subtype', 'Created on', 'Updated on',
  'Due date', 'Contractor', 'Discipline',
]
const workbookBytes = (rows) => {
  const workbook = XLSX.utils.book_new()
  const worksheet = XLSX.utils.aoa_to_sheet(rows)
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Issues')
  return XLSX.write(workbook, { type: 'array', bookType: 'xlsx', cellDates: true })
}
const currentBytes = workbookBytes([
  issueHeaders,
  ['BIM-100', 'Existing issue', 'Open', 'Coordination', '2026-06-01', '', '2026-06-10', 'Old Trade', 'BIM'],
])
const accHeaders = issueHeaders.map((header) => header === 'Subtype' ? 'Type' : header)
const accBytes = workbookBytes([
  [...accHeaders, 'Category', 'Created By', 'Created By (Company)'],
  ['BIM-100', 'Transferred existing issue', 'Closed', 'Access', '2026-07-01', '2026-07-02', '2026-07-05', '', 'BIM', 'Coordination', 'Peter Autodesk', 'Other'],
  ['BIM-101', 'New LotusWorks issue', 'Pending', 'Clearance', '2026-07-08', '', '2026-07-20', 'Trade B', 'Electrical', 'Field', 'Jamie Doe - LotusWorks', 'Other'],
  ['bim-101', 'Duplicate export row', 'Pending', 'Clearance', '2026-07-08', '', '2026-07-20', 'Trade B', 'Electrical', 'Field', 'LotusWorks', 'Other'],
  ['BIM-102', 'Other owner issue', 'Open', 'Quality', '2026-07-09', '', '', 'Trade C', 'Mechanical', 'Field', 'Outside Contractor', 'LotusWorks'],
  ['', 'Missing ID', 'Open', 'Quality', '2026-07-09', '', '', '', 'Mechanical', 'Field', 'LotusWorks', 'Other'],
])
const currentFile = new File([currentBytes], 'BIM_Issues_Log.xlsx')
const accFile = new File([accBytes], 'ACC_Issues_Export.xlsx')
const current = await prepareIssueWorkbook(currentFile, 'current')
const acc = await prepareIssueWorkbook(accFile, 'acc')
const analysis = reconcileIssueRows(current, acc)

if (analysis.trackedExistingIds !== 1) throw new Error('The current BIM log did not establish the tracked ID set.')
if (analysis.lotusWorksRows !== 2) throw new Error('Existing and new LotusWorks issue selection returned the wrong row count.')
if (analysis.updatedIssues.length !== 1 || analysis.updatedIssues[0].id !== 'BIM-100') {
  throw new Error('The transferred existing issue was not selected for update.')
}
if (analysis.skippedDuplicateIds !== 1) throw new Error('Duplicate ACC IDs were not skipped.')
if (analysis.skippedMissingIds !== 1) throw new Error('Rows without IDs were not skipped.')
if (analysis.excludedOtherOwners !== 1) throw new Error('Non-LotusWorks creators were not excluded.')
if (analysis.newIssues.length !== 1 || analysis.newIssues[0].id !== 'BIM-101') {
  throw new Error('The expected new LotusWorks issue was not selected.')
}

const output = buildUpdatedIssueWorkbook(current, analysis)
const updated = XLSX.read(output.bytes, { type: 'array', cellDates: true })
const rows = XLSX.utils.sheet_to_json(updated.Sheets.Issues, { defval: '', raw: false })
if (rows.length !== 2) throw new Error('The updated workbook should contain one appended row.')
const updatedExisting = rows[0]
const appended = rows[1]
if (updatedExisting.Status !== 'Closed' || updatedExisting.Title !== 'Transferred existing issue') {
  throw new Error('The existing issue was not updated from the ACC export.')
}
if (updatedExisting.Contractor !== 'Old Trade') {
  throw new Error('A blank ACC value overwrote an existing BIM log value.')
}
if (appended.ID !== 'BIM-101' || appended.Status !== 'Pending' || appended.Contractor !== 'Trade B') {
  throw new Error('The appended issue fields do not match the ACC export.')
}
if (updatedExisting.Subtype !== 'Access' || appended.Subtype !== 'Clearance') {
  throw new Error('ACC Type was not mapped into the BIM Subtype column.')
}
if (!output.fileName.includes('BIM_Issues_Log-Updated-')) {
  throw new Error('The updated workbook filename is not traceable to the source log.')
}
if (current.rows.length !== 1) throw new Error('The source workbook model was mutated.')
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
  console.log('Manual issue update refreshes tracked IDs, adds new LotusWorks IDs, preserves nonblank history, and writes the updated workbook.')
} finally {
  await rm(temporaryDirectory, { recursive: true, force: true })
}
