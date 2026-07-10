export interface ChangelogEntry {
  version: string
  date: string
  type: 'feature' | 'fix' | 'major'
  notes: string[]
}

export const CHANGELOG: ChangelogEntry[] = [
  {
    version: '1.0.4',
    date: '2026-07-10T13:46:30Z',
    type: 'feature',
    notes: [
      'Weekly Smartsheet exports can now be dropped directly into the standalone HTML app as XLS, XLSX, or CSV files',
      'The app automatically identifies all four required logs by filename and column headers without an API token or corporate proxy',
      'Imported spreadsheet data stays in the current browser session, and legacy saved Smartsheet credentials are removed automatically',
      'Report slides now retain their full 16:9 width at 1280px laptop viewports instead of collapsing into the file-panel column',
    ],
  },
  {
    version: '1.0.3',
    date: '2026-07-09T20:04:56Z',
    type: 'feature',
    notes: [
      'Charts now use smoother series, clearer scales, direct data labels, issue badges, and interactive hover details',
      'Slide cards, chart panels, and the BIM issue table now have consistent gutters and stay clear of the GC template rails',
      'The application header now keeps a consistent gap above the Smartsheet workspace and report filters',
    ],
  },
  {
    version: '1.0.2',
    date: '2026-07-09T19:00:00Z',
    type: 'fix',
    notes: [
      'Update modal now downloads the packaged HTML file directly instead of opening a new browser tab',
      'Download button and status copy now make it clear that the update is a single HTML app file',
    ],
  },
  {
    version: '1.0.1',
    date: '2026-07-09T18:00:00Z',
    type: 'fix',
    notes: [
      'App now opens on a clean Smartsheet connection screen instead of loading demo report data automatically',
      'Preview data is available only through the explicit Preview Layout button',
      'Report canvas and chart spacing were widened, and the sheet-mapping rail now collapses after a report loads so charts have more room',
    ],
  },
  {
    version: '1.0.0',
    date: '2026-07-09T12:00:00Z',
    type: 'major',
    notes: [
      'Initial QA/QC weekly report dashboard',
      'Smartsheet connection with saved auth and automatic sheet mapping',
      'OAC reporting cutoff, slicers, issue metrics, inspection metrics, welding metrics',
      'PowerPoint export with a reduced-height content band for GC slide templates',
    ],
  },
]
