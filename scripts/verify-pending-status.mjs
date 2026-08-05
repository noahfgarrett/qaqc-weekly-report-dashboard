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
import { buildReportModel, mergeFilters } from ${JSON.stringify(resolve(root, 'src/calculations/report.ts'))}

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
if (!report.issueTable.some((row) => row.id === 'BIM-100' && row.status === 'Pending')) {
  throw new Error('Pending is missing from BIM Issues Detail.')
}

const pendingOnly = buildReportModel(bundle, mergeFilters({ oac: true, statuses: ['Pending'] }), now)
const pendingMetric = (id) => pendingOnly.kpis.find((item) => item.id === id)?.rawValue
if (pendingMetric('total-opened') !== 1 || pendingMetric('total-closed') !== 0 || pendingMetric('remaining-open') !== 1) {
  throw new Error('The Pending Status filter does not preserve open-issue math.')
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
  console.log('Pending status counts as open and remains visible in filters and issue details.')
} finally {
  await rm(temporaryDirectory, { recursive: true, force: true })
}
