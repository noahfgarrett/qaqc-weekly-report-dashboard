import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import { build } from 'vite'

const root = process.cwd()
const temporaryDirectory = await mkdtemp(resolve(tmpdir(), 'qaqc-data-quality-'))
const entryPath = resolve(temporaryDirectory, 'data-quality-check.ts')
const outputDirectory = resolve(temporaryDirectory, 'dist')

const entrySource = `
import { applyDataQualityEdits, auditDataQuality } from ${JSON.stringify(resolve(root, 'src/services/dataQuality.ts'))}

const bundle = {
  source: 'files',
  sheets: {
    bimIssues: {
      id: 'bim',
      name: 'BIM Issues Log',
      rows: [
        {
          __rowNumber: 7,
          ID: '100',
          Status: 'Open',
          Type: 'Coordination',
          Title: 'Lotus issue',
          Contractor: '',
          Discipline: 'BIM',
          'Created On': '2026-08-24',
          'Due Date': '2026-09-04',
          'Created By': 'Samuel Leach LotusWorks',
        },
        {
          __rowNumber: 8,
          ID: '101',
          Status: 'Closed',
          Type: 'Quality',
          Title: 'Closed issue',
          Contractor: 'Turner',
          Discipline: 'Mechanical',
          'Created On': '2026-08-18',
          'Closed At': '',
          'Updated On': '',
          'Due Date': '2026-08-29',
          'Created By': 'Alex LotusWorks',
        },
        {
          __rowNumber: 9,
          ID: '999',
          Status: 'Open',
          Type: '',
          Title: '',
          Contractor: '',
          Discipline: '',
          'Created On': '',
          'Due Date': '',
          'Created By': 'Peter Autodesk',
        },
      ],
    },
    mechanical: {
      id: 'mechanical',
      name: 'Mechanical / Process Inspection Log',
      rows: [{ __rowNumber: 4, 'Inspection Phase': 'Final', 'Work Week Observed': "WW35'2026", 'General Contractor': '', Discipline: 'Mechanical', Subtype: 'Final' }],
    },
    electrical: {
      id: 'electrical',
      name: 'Electrical Inspection Log',
      rows: [{ __rowNumber: 5, 'Inspection Phase': '', 'Work Week Observed': "WW35'2026", 'General Contractor': '', Discipline: 'Electrical', Subtype: 'Final', 'Issue?': '' }],
    },
    welding: {
      id: 'welding',
      name: 'Welding Signoffs',
      rows: [{ __rowNumber: 11, NO: '', 'WELD WORK WEEK': '', Contractor: '', Discipline: '', Subtype: '', SIGNATURE: '', 'ISSUE CREATED? (Put BIM # If Yes)': '' }],
    },
  },
}

const audit = auditDataQuality(bundle)
const fields = audit.findings.map((finding) => finding.role + ':' + finding.rowIndex + ':' + finding.field)
const expected = [
  'bimIssues:0:contractor',
  'bimIssues:1:closedOn',
  'mechanical:0:contractor',
  'electrical:0:inspectionPhase',
  'electrical:0:contractor',
  'welding:0:weldNumber',
  'welding:0:weldWorkWeek',
  'welding:0:contractor',
  'welding:0:discipline',
  'welding:0:subtype',
]
if (fields.join('|') !== expected.join('|')) {
  throw new Error('Blank-field audit mismatch: ' + fields.join(', '))
}
if (audit.findings.some((finding) => finding.role === 'bimIssues' && finding.rowIndex === 2)) {
  throw new Error('A non-LotusWorks BIM issue was incorrectly included in the correction review.')
}
if (audit.findings.some((finding) => finding.field === 'signature' || finding.field === 'issueCreated')) {
  throw new Error('Intentional blank weld or issue indicators were incorrectly required.')
}
const contractorFindings = audit.findings.filter((finding) => finding.field === 'contractor')
if (contractorFindings.length !== 4 || contractorFindings.some((finding) => audit.initialValues[finding.id] !== 'Bechtel')) {
  throw new Error('Blank contractors were not consistently suggested as Bechtel.')
}

const values = { ...audit.initialValues }
audit.findings.forEach((finding) => {
  if (finding.field === 'closedOn') values[finding.id] = '2026-08-28'
  if (finding.field === 'inspectionPhase') values[finding.id] = 'Final'
  if (finding.field === 'weldNumber') values[finding.id] = 'W-204'
  if (finding.field === 'weldWorkWeek') values[finding.id] = "WW35'2026"
})
const corrected = applyDataQualityEdits(bundle, audit.findings, values)
if (corrected.sheets.bimIssues.rows[0].Contractor !== 'Bechtel') throw new Error('BIM Contractor correction was not applied.')
if (corrected.sheets.bimIssues.rows[1]['Closed At'] !== '2026-08-28') throw new Error('Closed At correction missed its source column.')
if (corrected.sheets.electrical.rows[0]['General Contractor'] !== 'Bechtel') throw new Error('General Contractor correction missed its source column.')
if (corrected.sheets.electrical.rows[0]['Inspection Phase'] !== 'Final') throw new Error('Inspection Phase correction was not applied.')
if (corrected.sheets.welding.rows[0].NO !== 'W-204' || corrected.sheets.welding.rows[0]['WELD WORK WEEK'] !== "WW35'2026") {
  throw new Error('Welding corrections were not applied.')
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
      rollupOptions: { output: { entryFileNames: 'data-quality-check.mjs', format: 'es' } },
    },
  })
  await import(`${pathToFileURL(resolve(outputDirectory, 'data-quality-check.mjs')).href}?t=${Date.now()}`)
  console.log('Missing report fields are audited, suggested, and applied without flagging intentional blanks.')
} finally {
  await rm(temporaryDirectory, { recursive: true, force: true })
}
