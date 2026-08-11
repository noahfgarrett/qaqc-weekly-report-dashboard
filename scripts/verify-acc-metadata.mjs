import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import { build } from 'vite'

const root = process.cwd()
const temporaryDirectory = await mkdtemp(resolve(tmpdir(), 'qaqc-acc-metadata-'))
const entryPath = resolve(temporaryDirectory, 'acc-metadata-check.ts')
const outputDirectory = resolve(temporaryDirectory, 'dist')

const entrySource = `
import { buildReportModel, mergeFilters } from ${JSON.stringify(resolve(root, 'src/calculations/report.ts'))}

const excelSerial = (isoDate) => {
  const [year, month, day] = isoDate.split('-').map(Number)
  return (Date.UTC(year, month - 1, day) - Date.UTC(1899, 11, 30)) / 86400000
}
const emptySheet = (name) => ({ id: name, name, rows: [] })
const bundle = {
  source: 'files',
  sheets: {
    bimIssues: {
      id: 'acc-metadata',
      name: 'ACC Issues Export',
      rows: [
        {
          ID: '3001', Status: 'Pending', Type: 'Coordination', Subtype: 'Wrong subtype',
          'Created On': '2026-07-24', 'Updated On': '2026-08-06',
          'bim360 created_by': 'Samuel Leach LOTUSWORKS',
          'BIM360 Created On': excelSerial('2026-08-04'), 'BIM360 Closed On': '',
        },
        {
          ID: '3002', Status: 'Closed', Type: 'Access',
          'Created On': '2026-07-24', 'Updated On': '2026-07-24',
          'BIM360_Created By': 'LotusWorks', 'BIM360_Created On': '2026-08-05',
          'BIM360_Closed On': '2026-08-07',
        },
        {
          ID: '3003', Status: 'Closed', Type: 'Quality',
          'Created On': '2026-07-24', 'Updated On': excelSerial('2026-08-06'),
          'BIM360_Created By': 'LotusWorks', 'BIM360_Created On': '2026-07-01',
          'BIM360_Closed On': '',
        },
        {
          ID: '3004', Status: 'Closed', Type: 'Safety',
          'Created On': '2026-07-24', 'Updated On': '2026-07-24',
          'BIM360_Created By': 'LotusWorks', 'BIM360_Created On': '2026-07-01',
          'BIM360_Closed On': '08/10/2026',
        },
        {
          ID: '3005', Status: 'Open', Type: 'Field',
          'Created On': '2026-07-24', 'Updated On': '2026-08-06',
          'BIM360_Created By': 'LotusWorks', 'BIM360_Created On': '2026-07-01',
          'BIM360_Closed On': '',
        },
        {
          ID: '3006', Status: 'Pending', Type: '', Subtype: 'Fallback subtype',
          'Created On': new Date(2026, 7, 8, 12), 'Updated On': '2026-08-08',
          'Created By': 'SAM LOTUSWORKS',
          'BIM360_Created By': '', 'BIM360_Created On': '', 'BIM360_Closed On': '',
        },
        {
          ID: '3007', Status: 'Open', Type: 'Outside fallback',
          'Created On': '2026-08-05', 'Updated On': '2026-08-06',
          'Created By': 'Outside Contractor',
          'BIM360_Created By': '', 'BIM360_Created On': '', 'BIM360_Closed On': '',
        },
        {
          ID: '3008', Status: 'Open', Type: 'Outside legacy',
          'Created On': '2026-08-05', 'Updated On': '2026-08-06', 'Created By': 'LotusWorks',
          'BIM360_Created By': 'Outside Contractor', 'BIM360_Created On': '2026-08-05',
          'BIM360_Closed On': '',
        },
        {
          ID: '3009', Status: 'Void', Type: 'Void type',
          'Created On': '2026-08-05', 'Updated On': '2026-08-06',
          'BIM360_Created By': 'LotusWorks', 'BIM360_Created On': '2026-08-05',
          'BIM360_Closed On': '',
        },
        {
          ID: '3010', Status: 'Complete', Type: 'Completion alias',
          'Created On': '2026-07-24', 'Updated On': '2026-07-24',
          'BIM360_Created By': 'LotusWorks', 'BIM360_Created On': '2026-07-02',
          'BIM360_Closed On': excelSerial('2026-08-09'),
        },
        {
          ID: '3012', Status: 'Open', Type: 'RFI',
          'Created On': '2026-08-09', 'Updated On': '2026-08-09', 'Issue Owner': 'LotusWorks',
          'BIM360_Created By': '', 'BIM360_Created On': '', 'BIM360_Closed On': '',
        },
      ],
    },
    mechanical: emptySheet('Mechanical / Process Inspection Log'),
    electrical: emptySheet('Electrical Inspection Log'),
    welding: emptySheet('Welding Signoffs by Work Week'),
  },
}

const report = buildReportModel(bundle, mergeFilters({ oac: true }), new Date(2026, 7, 11, 12))
if (report.currentWeek.label !== "WW33'2026" || report.reportWeek.label !== "WW32'2026") {
  throw new Error('ACC metadata test is not using the expected current and reporting weeks.')
}

const metric = (id) => report.kpis.find((item) => item.id === id)?.rawValue
const expectedMetrics = new Map([
  ['total-opened', 8],
  ['total-closed', 3],
  ['opened-week', 4],
  ['closed-week', 3],
  ['remaining-open', 5],
])
expectedMetrics.forEach((expected, id) => {
  if (metric(id) !== expected) {
    throw new Error(id + ' expected ' + expected + ' but received ' + metric(id) + '.')
  }
})

const expectedDetails = new Map([
  ['3001', ['Opened in Report Week', 'Open', 'Coordination', "WW32'2026", 'Open']],
  ['3002', ['Opened + Closed in Report Week', 'Closed', 'Access', "WW32'2026", "WW32'2026"]],
  ['3003', ['Closed in Report Week', 'Closed', 'Quality', "WW27'2026", "WW32'2026"]],
  ['3004', ['Closed This Week', 'Closed', 'Safety', "WW27'2026", "WW33'2026"]],
  ['3006', ['Opened in Report Week', 'Open', 'Fallback subtype', "WW32'2026", 'Open']],
  ['3010', ['Closed in Report Week', 'Complete', 'Completion alias', "WW27'2026", "WW32'2026"]],
  ['3012', ['Opened in Report Week', 'Open', 'RFI', "WW32'2026", 'Open']],
])
if (report.issueTable.map((row) => row.id).join(',') !== '3012,3010,3006,3004,3003,3002,3001') {
  throw new Error('ACC metadata rows were not included, excluded, or sorted correctly.')
}
report.issueTable.forEach((row) => {
  const expected = expectedDetails.get(row.id)
  if (!expected) throw new Error(row.id + ' unexpectedly appeared in BIM Issues Detail.')
  const actual = [row.group, row.status, row.subtype, row.workWeek, row.workWeekClosed]
  if (actual.join('|') !== expected.join('|')) {
    throw new Error(row.id + ' did not populate its report metadata correctly: ' + actual.join('|'))
  }
})

if (report.filterOptions.subtypes.includes('Wrong subtype')) {
  throw new Error('Subtype overrode a populated ACC Type value.')
}
if (!report.filterOptions.subtypes.includes('Fallback subtype')) {
  throw new Error('A blank ACC Type did not fall back to Subtype.')
}
if (report.filterOptions.statuses.includes('Void')) {
  throw new Error('Void leaked into the status filter options.')
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
      rollupOptions: { output: { entryFileNames: 'acc-metadata-check.mjs', format: 'es' } },
    },
  })
  await import(`${pathToFileURL(resolve(outputDirectory, 'acc-metadata-check.mjs')).href}?t=${Date.now()}`)
  console.log('ACC ownership, legacy dates, Type metadata, statuses, work weeks, metrics, and detail groups populate correctly.')
} finally {
  await rm(temporaryDirectory, { recursive: true, force: true })
}
