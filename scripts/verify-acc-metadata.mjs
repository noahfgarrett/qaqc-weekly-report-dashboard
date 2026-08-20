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
          'Created On': '2026-07-24', 'Closed at': '2026-08-11', 'Updated On': '2026-07-24',
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
          ID: '3011', Status: 'Closed', Type: 'Closure priority',
          'Created On': '2026-07-01', 'Closed at': '2026-08-07', 'Updated On': '2026-08-11',
          'Created By': 'LotusWorks',
          'BIM360_Created By': '', 'BIM360_Created On': '', 'BIM360_Closed On': '',
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
  ['total-opened', 9],
  ['total-closed', 4],
  ['opened-week', 4],
  ['closed-week', 4],
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
  ['3011', ['Closed in Report Week', 'Closed', 'Closure priority', "WW27'2026", "WW32'2026"]],
  ['3012', ['Opened in Report Week', 'Open', 'RFI', "WW32'2026", 'Open']],
])
if (report.issueTable.map((row) => row.id).join(',') !== '3012,3011,3010,3006,3004,3003,3002,3001') {
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
if (report.issueTable.find((row) => row.id === '3011')?.workWeekClosed !== "WW32'2026") {
  throw new Error('Closed at did not take priority over an Updated On value from a later work week.')
}
if (report.issueTable.find((row) => row.id === '3002')?.workWeekClosed !== "WW32'2026") {
  throw new Error('A historical BIM closure date did not take priority over ACC Closed at and Updated On.')
}
if (report.issueTable.find((row) => row.id === '3003')?.workWeekClosed !== "WW32'2026") {
  throw new Error('A blank dedicated closure date did not fall back to Updated On.')
}
if (report.filterOptions.statuses.includes('Void')) {
  throw new Error('Void leaked into the status filter options.')
}

const closureMetric = (model) => model.kpis.find((item) => item.id === 'closure-rate')
if (closureMetric(report)?.tone !== 'good') {
  throw new Error('An improving Closure Rate must use the green good tone.')
}

const priorWeekIssues = [
  { ID: '4001', Status: 'Closed', 'Created On': '2026-07-01', 'Updated On': '2026-07-10' },
  { ID: '4002', Status: 'Closed', 'Created On': '2026-07-02', 'Updated On': '2026-07-11' },
  { ID: '4003', Status: 'Open', 'Created On': '2026-07-03', 'Updated On': '2026-07-03' },
  { ID: '4004', Status: 'Open', 'Created On': '2026-07-04', 'Updated On': '2026-07-04' },
]
const newOpenIssues = [
  { ID: '4005', Status: 'Open', 'Created On': '2026-08-04', 'Updated On': '2026-08-04' },
  { ID: '4006', Status: 'Open', 'Created On': '2026-08-05', 'Updated On': '2026-08-05' },
  { ID: '4007', Status: 'Open', 'Created On': '2026-08-06', 'Updated On': '2026-08-06' },
  { ID: '4008', Status: 'Open', 'Created On': '2026-08-07', 'Updated On': '2026-08-07' },
]
const reportForIssues = (rows) => buildReportModel({
  ...bundle,
  sheets: {
    ...bundle.sheets,
    bimIssues: { id: 'closure-tone', name: 'ACC Issues Export', rows },
  },
}, mergeFilters({ oac: true }), new Date(2026, 7, 11, 12))

const decliningClosure = closureMetric(reportForIssues([...priorWeekIssues, ...newOpenIssues]))
if (decliningClosure?.rawValue !== 25 || decliningClosure.tone !== 'bad') {
  throw new Error('A declining Closure Rate must use the red bad tone.')
}
const unchangedClosure = closureMetric(reportForIssues(priorWeekIssues))
if (unchangedClosure?.rawValue !== 50 || unchangedClosure.tone !== 'neutral') {
  throw new Error('An unchanged Closure Rate must use the neutral tone.')
}

const contractorFallbackBundle = {
  ...bundle,
  sheets: {
    ...bundle.sheets,
    bimIssues: {
      id: 'contractor-fallback',
      name: 'ACC Issues Export',
      rows: [
        { ID: '#1021', Status: 'Open', 'Created On': '2026-08-05', 'Created By': 'Stuart Barrett LotusWorks', Contractor: '', Discipline: '' },
        { ID: 'BIM-1021', Status: 'Open', 'Created On': '2026-08-05', 'Created By': 'Stuart Barrett LotusWorks', Contractor: 'Future Contractor', Discipline: '' },
      ],
    },
  },
}
const contractorFallbackReport = buildReportModel(
  contractorFallbackBundle,
  mergeFilters({ oac: true }),
  new Date(2026, 7, 11, 12),
)
if (!contractorFallbackReport.filterOptions.contractors.includes('Bechtel')) {
  throw new Error('Blank Contractor on issue #1021 did not fall back to Bechtel.')
}
if (!contractorFallbackReport.filterOptions.contractors.includes('Future Contractor')) {
  throw new Error('A populated Contractor on issue #1021 was overwritten by the fallback.')
}
const bechtelOnly = buildReportModel(
  contractorFallbackBundle,
  mergeFilters({ oac: true, contractors: ['Bechtel'] }),
  new Date(2026, 7, 11, 12),
)
if (bechtelOnly.kpis.find((item) => item.id === 'total-opened')?.rawValue !== 1) {
  throw new Error('The #1021 Bechtel fallback did not participate in Contractor filtering.')
}
if (!contractorFallbackReport.filterOptions.disciplines.includes('Unassigned')) {
  throw new Error('Blank issue Discipline no longer remains available as Unassigned.')
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
  console.log('ACC metadata, report groups, metrics, and directional Closure Rate tones populate correctly.')
} finally {
  await rm(temporaryDirectory, { recursive: true, force: true })
}
