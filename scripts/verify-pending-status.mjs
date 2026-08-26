import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import { build } from 'vite'

const root = process.cwd()
const temporaryDirectory = await mkdtemp(resolve(tmpdir(), 'qaqc-pending-status-'))
const entryPath = resolve(temporaryDirectory, 'pending-status-check.ts')
const outputDirectory = resolve(temporaryDirectory, 'dist')

const entrySource = `
import { buildReportModel, mergeFilters, selectIssueDetailExportRows } from ${JSON.stringify(resolve(root, 'src/calculations/report.ts'))}

const emptySheet = (name) => ({ id: name, name, rows: [] })
const bundle = {
  source: 'files',
  sheets: {
    bimIssues: {
      id: 'bim',
      name: 'BIM Issues Log',
      rows: [
        { ID: 'BIM-100', Status: ' pending ', 'Created On': '2026-07-01', 'Updated On': '2026-07-01' },
        { ID: 'BIM-101', Status: 'Closed', 'Created On': '2026-07-01', 'Updated On': '2026-07-02' },
        { ID: 'BIM-102', Status: 'Void', 'Created On': '2026-07-01', 'Updated On': '2026-07-01' },
      ],
    },
    mechanical: emptySheet('Mechanical / Process Inspection Log'),
    electrical: emptySheet('Electrical Inspection Log'),
    welding: emptySheet('Welding Signoffs by Work Week'),
  },
}
const now = new Date(2026, 6, 9, 12)
const report = buildReportModel(bundle, mergeFilters({ oac: true }), now)
const metric = (id) => report.kpis.find((item) => item.id === id)?.rawValue

if (!report.filterOptions.statuses.includes('Pending')) throw new Error('Pending is missing from the Status slicer.')
if (metric('total-opened') !== 2) throw new Error('Pending must count toward Total Issues Opened.')
if (metric('total-closed') !== 1) throw new Error('Pending must not count toward Total Issues Closed.')
if (metric('remaining-open') !== 1) throw new Error('Pending must count toward Issues Remaining Open.')
if (!report.issueTable.some((row) => row.id === 'BIM-100' && row.status === 'Open')) {
  throw new Error('Pending is not displayed as Open in BIM Issues Detail.')
}

const pendingOnly = buildReportModel(bundle, mergeFilters({ oac: true, statuses: ['Pending'] }), now)
const pendingMetric = (id) => pendingOnly.kpis.find((item) => item.id === id)?.rawValue
if (pendingMetric('total-opened') !== 1 || pendingMetric('total-closed') !== 0 || pendingMetric('remaining-open') !== 1) {
  throw new Error('The Pending Status filter does not preserve open-issue math.')
}

const strictDetailBundle = {
  ...bundle,
  sheets: {
    ...bundle.sheets,
    bimIssues: {
      id: 'bim-strict-detail',
      name: 'BIM Issues Log',
      rows: [
        { ID: '1018', Status: 'Pending', 'Created On': '2026-07-24', 'Updated On': '2026-07-31', 'BIM360_Created By': 'LotusWorks', 'BIM360_Created On': '2026-07-24', 'BIM360_Closed On': '' },
        { ID: '1019', Status: 'Pending', Type: 'Clearance', Subtype: 'Wrong legacy subtype', 'Created On': '2026-07-24', 'Updated On': '2026-07-31', 'BIM360_Created By': 'LotusWorks', 'BIM360_Created On': '2026-07-29', 'BIM360_Closed On': '' },
        { ID: '1020', Status: 'Closed', 'Created On': '2026-07-24', 'Updated On': '2026-07-24', 'BIM360_Created By': 'LotusWorks', 'BIM360_Created On': '2026-07-01', 'BIM360_Closed On': '2026-07-31' },
        { ID: '1021', Status: 'Closed', 'Created On': '2026-07-24', 'Updated On': '2026-07-24', 'BIM360_Created By': 'LotusWorks', 'BIM360_Created On': '2026-07-29', 'BIM360_Closed On': '2026-07-31' },
        { ID: '1022', Status: 'Closed', 'Created On': '2026-07-24', 'Updated On': '2026-08-04', 'BIM360_Created By': 'LotusWorks', 'BIM360_Created On': '2026-07-01', 'BIM360_Closed On': '' },
        { ID: '1023', Status: 'Pending', 'Created On': '2026-07-30', 'Updated On': '2026-07-31', 'Created By': 'LotusWorks', 'BIM360_Created By': '', 'BIM360_Created On': '', 'BIM360_Closed On': '' },
        { ID: '2000', Status: 'Open', 'Created On': '2026-07-29', 'Updated On': '2026-07-31', 'Created By': 'LotusWorks', 'BIM360_Created By': 'Outside Contractor', 'BIM360_Created On': '2026-07-29', 'BIM360_Closed On': '' },
        { ID: '2001', Status: 'Open', 'Created On': '2026-07-30', 'Updated On': '2026-07-31', 'Created By': 'Outside Contractor', 'BIM360_Created By': '', 'BIM360_Created On': '', 'BIM360_Closed On': '' },
        { ID: '999', Status: 'Open', 'Created On': '2026-07-24', 'Updated On': '2026-07-31', 'BIM360_Created By': 'LotusWorks', 'BIM360_Created On': '2026-04-24', 'BIM360_Closed On': '' },
      ],
    },
  },
}
const strictDetailReport = buildReportModel(
  strictDetailBundle,
  mergeFilters({ oac: true }),
  new Date(2026, 7, 6, 12),
)
if (strictDetailReport.reportWeek.label !== "WW31'2026") {
  throw new Error('The strict issue-detail regression is not using the expected reporting week.')
}
const strictIds = strictDetailReport.issueTable.map((row) => row.id)
if (strictIds.join(',') !== '1023,1022,1021,1020,1019,1018,999') {
  throw new Error('BIM Issues Detail did not keep report activity first and append supplemental open issues.')
}
const expectedGroups = new Map([
  ['1019', 'Opened in Report Week'],
  ['1020', 'Closed in Report Week'],
  ['1021', 'Opened + Closed in Report Week'],
  ['1022', 'Closed This Week'],
  ['1023', 'Opened in Report Week'],
  ['1018', 'Open for Discussion'],
  ['999', 'Open for Discussion'],
])
strictDetailReport.issueTable.forEach((row) => {
  if (row.group !== expectedGroups.get(row.id)) {
    throw new Error(row.id + ' was assigned to the wrong issue-detail card.')
  }
})
if (strictDetailReport.issueTable.find((row) => row.id === '1019')?.status !== 'Open') {
  throw new Error('A reporting-week Pending issue must display as Open.')
}
if (strictDetailReport.issueTable.find((row) => row.id === '1019')?.subtype !== 'Clearance') {
  throw new Error('ACC Type did not take priority over the legacy Subtype field.')
}
if (strictDetailReport.issueTable.some((row) => row.id === '2000' || row.id === '2001')) {
  throw new Error('Legacy and fallback ACC ownership filters did not exclude outside issues.')
}
const exportFixture = [
  ...Array.from({ length: 5 }, (_, index) => ({ ...strictDetailReport.issueTable[0], id: 'R-' + index, group: 'Opened in Report Week' })),
  ...Array.from({ length: 20 }, (_, index) => ({ ...strictDetailReport.issueTable.at(-1), id: 'D-' + index, group: 'Open for Discussion' })),
]
const exportRows = selectIssueDetailExportRows(exportFixture, 14)
if (exportRows.length !== 14 || exportRows.filter((row) => row.group === 'Open for Discussion').length !== 9) {
  throw new Error('Supplemental open issues did not stop after filling the unused export-page rows.')
}
const strictMetric = (id) => strictDetailReport.kpis.find((item) => item.id === id)?.rawValue
if (strictMetric('total-opened') !== 7 || strictMetric('total-closed') !== 2 || strictMetric('remaining-open') !== 5) {
  throw new Error('Legacy ownership and date priority did not propagate through dashboard metrics.')
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
      outDir: outputDirectory,
      emptyOutDir: true,
      rollupOptions: { output: { entryFileNames: 'pending-status-check.mjs', format: 'es' } },
    },
  })
  await import(`${pathToFileURL(resolve(outputDirectory, 'pending-status-check.mjs')).href}?t=${Date.now()}`)
  console.log('Pending counts as open and supplemental open issues follow the report activity rows.')
} finally {
  await rm(temporaryDirectory, { recursive: true, force: true })
}
