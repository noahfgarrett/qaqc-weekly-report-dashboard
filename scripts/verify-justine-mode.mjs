import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import { build } from 'vite'

const root = process.cwd()
const temporaryDirectory = await mkdtemp(resolve(tmpdir(), 'qaqc-justine-mode-'))
const entryPath = resolve(temporaryDirectory, 'justine-mode-check.ts')
const outputDirectory = resolve(temporaryDirectory, 'dist')

const entrySource = `
import { buildReportModel, mergeFilters } from ${JSON.stringify(resolve(root, 'src/calculations/report.ts'))}

const bundle = {
  source: 'files',
  sheets: {
    bimIssues: {
      id: 'bim',
      name: 'ACC Issues Export',
      rows: [
        { ID: '100', Status: 'Open', 'Created On': '2026-07-21' },
        { ID: '101', Status: 'Closed', 'Created On': '2026-07-28', 'Updated On': '2026-07-30' },
        { ID: '102', Status: 'Open', 'Created On': '2026-08-04' },
        { ID: '103', Status: 'Closed', 'Created On': '2026-08-05', 'Updated On': '2026-08-07' },
        { ID: '104', Status: 'Pending', 'Created On': '2026-08-10' },
        { ID: '105', Status: 'Closed', 'Created On': '2026-07-01', 'Updated On': '2026-08-11' },
      ],
    },
    mechanical: {
      id: 'mechanical',
      name: 'Mechanical / Process Inspection Log',
      rows: [
        { 'Inspection Phase': 'Final', 'Work Week Observed': "WW30'2026" },
        { 'Inspection Phase': 'SOR', 'Work Week Observed': "WW32'2026" },
        { 'Inspection Phase': 'Final', 'Work Week Observed': "WW33'2026" },
      ],
    },
    electrical: {
      id: 'electrical',
      name: 'Electrical Inspection Log',
      rows: [
        { 'Inspection Phase': 'Final', 'Work Week Observed': "WW31'2026", 'Issue?': 'No Issue Found' },
        { 'Inspection Phase': 'Final', 'Work Week Observed': "WW32'2026", 'Issue?': 'BIM-22' },
        { 'Inspection Phase': 'Final', 'Work Week Observed': "WW33'2026", 'Issue?': 'No Issue Found' },
        { 'Inspection Phase': 'Draft', 'Work Week Observed': "WW33'2026", 'Issue?': 'BIM-23' },
      ],
    },
    welding: {
      id: 'welding',
      name: 'Welding Signoffs by Work Week',
      rows: [
        { NO: 'W-1', 'WELD WORK WEEK': "WW30'2026", SIGNATURE: 'Signed' },
        { NO: 'W-2', 'WELD WORK WEEK': "WW31'2026", SIGNATURE: '' },
        { NO: 'W-3', 'WELD WORK WEEK': "WW31'2026", SIGNATURE: '' },
        { NO: 'W-4', 'WELD WORK WEEK': "WW32'2026", SIGNATURE: 'Signed' },
        { NO: 'W-5', 'WELD WORK WEEK': "WW32'2026", SIGNATURE: '' },
        { NO: 'W-6', 'WELD WORK WEEK': "WW33'2026", SIGNATURE: 'Signed' },
      ],
    },
  },
}

const now = new Date(2026, 7, 11, 12)
const metric = (report, id) => report.kpis.find((item) => item.id === id)?.rawValue

const legacyOac = mergeFilters({ oac: true })
const legacyManual = mergeFilters({ oac: false })
if (legacyOac.reportingMode !== 'oac' || legacyManual.reportingMode !== 'manual') {
  throw new Error('Saved pre-Justine filter settings did not migrate correctly.')
}

const report = buildReportModel(bundle, mergeFilters({ reportingMode: 'justine' }), now)
if (report.reportingMode !== 'justine' || report.periodStartWeek.label !== "WW32'2026" || report.periodEndWeek.label !== "WW33'2026") {
  throw new Error('Justine did not resolve the previous-plus-current reporting period.')
}
if (report.periodLabel !== "WW32'2026 + WW33'2026" || report.cutoffDate.getTime() !== now.getTime()) {
  throw new Error('Justine period labeling or live cutoff is incorrect.')
}

const expectedMetrics = new Map([
  ['total-opened', 6],
  ['total-closed', 3],
  ['remaining-open', 3],
  ['opened-week', 3],
  ['closed-week', 2],
  ['inspections', 3],
  ['sors', 1],
  ['closure-rate', 50],
])
expectedMetrics.forEach((expected, id) => {
  if (metric(report, id) !== expected) throw new Error(id + ' did not use the Justine reporting period.')
})

const expectedDeltas = new Map([
  ['opened-week', 1],
  ['closed-week', 1],
  ['inspections', 1],
])
expectedDeltas.forEach((expected, id) => {
  if (report.kpis.find((item) => item.id === id)?.delta !== expected) {
    throw new Error(id + ' did not compare against the preceding two-week period.')
  }
})

if (report.summary.electricalFinals !== 2 || report.summary.electricalIssuesFound !== 1) {
  throw new Error('Electrical summary cards did not aggregate both Justine weeks.')
}
if (report.summary.weldsChecked !== 3 || report.summary.weldsSigned !== 2 || Math.abs(report.summary.reportWeekSignoffRate - 66.6667) > 0.01) {
  throw new Error('Welding summary cards did not aggregate both Justine weeks.')
}
if (report.issueTable.map((row) => row.id).join(',') !== '105,104,103,102') {
  throw new Error('BIM Issues Detail did not use the Justine two-week activity window.')
}
if (report.issueTable.find((row) => row.id === '104')?.status !== 'Open') {
  throw new Error('Pending did not remain an open issue in Justine mode.')
}

const oac = buildReportModel(bundle, mergeFilters({ reportingMode: 'oac' }), now)
if (oac.periodLabel !== "WW32'2026" || metric(oac, 'opened-week') !== 2 || metric(oac, 'closed-week') !== 1) {
  throw new Error('OAC no longer uses the previous completed work week.')
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
      rollupOptions: { output: { entryFileNames: 'justine-mode-check.mjs', format: 'es' } },
    },
  })
  await import(`${pathToFileURL(resolve(outputDirectory, 'justine-mode-check.mjs')).href}?t=${Date.now()}`)
  console.log('Justine mode aggregates two live work weeks while OAC and saved settings remain compatible.')
} finally {
  await rm(temporaryDirectory, { recursive: true, force: true })
}
