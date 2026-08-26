import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import { build } from 'vite'

const root = process.cwd()
const temporaryDirectory = await mkdtemp(resolve(tmpdir(), 'qaqc-open-aging-'))
const entryPath = resolve(temporaryDirectory, 'open-aging-check.ts')
const outputDirectory = resolve(temporaryDirectory, 'dist')

const entrySource = `
import { buildReportModel, mergeFilters } from ${JSON.stringify(resolve(root, 'src/calculations/report.ts'))}

const emptySheet = (name) => ({ id: name, name, rows: [] })
const rows = [
  { ID: '5001', Status: 'Open', 'Created On': '2026-08-16' },
  { ID: '5002', Status: 'Pending', 'Created On': '2026-08-10' },
  { ID: '5003', Status: 'Open', 'Created On': '2026-08-03' },
  { ID: '5004', Status: 'Pending', 'Created On': '2026-07-27' },
  { ID: '5005', Status: 'Open', 'Created On': '2026-07-20' },
  { ID: '5006', Status: 'Closed', 'Created On': '2026-07-01', 'Closed At': '2026-07-10' },
  { ID: '5007', Status: 'Void', 'Created On': '2026-07-01' },
  { ID: '5008', Status: 'Open', 'Created On': '2026-08-17' },
]
const bundle = {
  source: 'files',
  sheets: {
    bimIssues: { id: 'open-aging', name: 'ACC Issues Export', rows },
    mechanical: emptySheet('Mechanical / Process Inspection Log'),
    electrical: emptySheet('Electrical Inspection Log'),
    welding: emptySheet('Welding Signoffs by Work Week'),
  },
}

const report = buildReportModel(bundle, mergeFilters({ reportingMode: 'oac', oac: true }), new Date(2026, 7, 17, 12))
if (report.currentWeek.label !== "WW34'2026" || report.reportWeek.label !== "WW33'2026") {
  throw new Error('Open aging regression is not using the expected work weeks.')
}

const labels = report.openAging.map((bucket) => bucket.label)
const counts = report.openAging.map((bucket) => bucket.count)
if (labels.join(',') !== '0-1,1-2,2-3,3-4,4+') {
  throw new Error('Open aging bucket labels changed: ' + labels.join(','))
}
if (counts.join(',') !== '1,1,1,1,1') {
  throw new Error('Open and Pending issue ages did not land in the five weekly buckets: ' + counts.join(','))
}
const colors = report.openAging.map((bucket) => bucket.color)
if (colors.join(',') !== '#B8C4D0,#9EAFBF,#778C9F,#5D7185,#3F5367') {
  throw new Error('Open aging bars must darken with issue age: ' + colors.join(','))
}
const overallColors = report.aging.map((bucket) => bucket.color)
if (overallColors.join(',') !== '#B8C4D0,#778C9F,#3F5367') {
  throw new Error('Overall aging bars must use the matching slate age scale: ' + overallColors.join(','))
}
if (report.openAging.reduce((sum, bucket) => sum + bucket.count, 0) !== 5) {
  throw new Error('Closed, Void, or current-week issues leaked into OAC open aging.')
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
      rollupOptions: { output: { entryFileNames: 'open-aging-check.mjs', format: 'es' } },
    },
  })
  await import(`${pathToFileURL(resolve(outputDirectory, 'open-aging-check.mjs')).href}?t=${Date.now()}`)
  console.log('Open and Pending issues populate the five weekly aging buckets correctly.')
} finally {
  await rm(temporaryDirectory, { recursive: true, force: true })
}
