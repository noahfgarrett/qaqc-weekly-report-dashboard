export interface ChangelogEntry {
  version: string
  date: string
  type: 'feature' | 'fix' | 'major'
  notes: string[]
}

export const CHANGELOG: ChangelogEntry[] = [
  {
    version: '1.0.11',
    date: '2026-07-15T13:00:55Z',
    type: 'feature',
    notes: [
      'Electrical and Mechanical inspection rows now use General Contractor for Contractor filtering, while Welding Signoffs ignore the Contractor slicer',
      'Welding Signoffs now combines signed and total weld bars with a secondary-axis sign-off percentage line, a 10% baseline, direct labels, and issue-count badges above each bar',
      'Electrical inspections now label every weekly value and position issue-count badges below the corresponding line point',
      'The sign-off summary now separates the average across active welding weeks from the selected report-week percentage',
      'Issue details now sort from highest natural ID to lowest across the dashboard, PDF, and PowerPoint continuation pages',
      'Project-to-date issue cards were reordered and clarified, issue aging now uses green, orange, and red severity colors, and report surfaces use more consistent icons and subtle shadows',
      'Week-over-week comparisons now spell out Previous Week instead of using the PW abbreviation',
    ],
  },
  {
    version: '1.0.9',
    date: '2026-07-15T11:23:09Z',
    type: 'major',
    notes: [
      'The overview report was rebuilt around equal Project to Date and reporting-week KPI panels, with the redundant title bar removed to give every chart more room',
      'Issues by Work Week now spans the full slide with Opened and Closed bars, a secondary-axis Remaining Open line, readable data labels, and compact WW## axis labels',
      'Cumulative Opened vs Closed and Issue Aging now use the full lower report band with clearer axes and direct values',
      'BIM issue details now continue across as many PDF and PowerPoint slides as needed instead of stopping after the first 16 rows',
      'PowerPoint export now downloads reliably from the standalone HTML app, and Electrical and Welding imports accept common Smartsheet header and work-week variations',
      'PowerPoint charts remain editable vector objects when resized, while PDF slides now render at 3x resolution for sharper report distribution',
    ],
  },
  {
    version: '1.0.8',
    date: '2026-07-13T20:03:00Z',
    type: 'feature',
    notes: [
      'Issues by Work Week is now two stacked charts sharing the work-week axis: the Remaining Open backlog line on top and the weekly Opened and Closed bars below, each on its own scale',
      'Weekly opened and closed bars are now clearly readable instead of being flattened beneath the much larger backlog line',
      'The exported PPTX and PDF deck use the same stacked layout',
    ],
  },
  {
    version: '1.0.7',
    date: '2026-07-13T19:52:14Z',
    type: 'feature',
    notes: [
      'The overview now leads with the charts: the KPI band was compacted into a single row per group so the trend, cumulative, and aging charts take significantly more of the slide',
      'The Cumulative Opened vs Closed chart is now full size and legible instead of a cramped thumbnail',
      'Removed a redundant context-pill row and tightened chart panels to give every plot more room',
    ],
  },
  {
    version: '1.0.6',
    date: '2026-07-13T12:58:54Z',
    type: 'major',
    notes: [
      'The dashboard and exported deck share a new refined, professional design built around a single slate-blue accent, with red, amber, and green reserved for status',
      'KPI metrics are now grouped into project-to-date and current-week clusters, led by a Remaining Open headline with real multi-week trend sparklines',
      'Every chart was rebuilt for clarity: consistent color meaning for opened, closed, and remaining issues, clean integer axes, and colorblind-safe palettes',
      'The welding chart now shows sign-off rate on a single axis, reading signed welds within the total-weld track instead of a second scale',
      'Exported PPTX and PDF decks match the new visual language for a consistent hand-off to the General Contractor',
    ],
  },
  {
    version: '1.0.5',
    date: '2026-07-13T11:19:48Z',
    type: 'feature',
    notes: [
      'A single ZIP containing all weekly Smartsheet exports can now generate the complete report',
      'Full folders can be selected or dragged into the app with recursive support for nested directories',
      'ZIP imports ignore unrelated files and macOS metadata while preserving automatic four-log mapping',
      'Archive size limits prevent unusually large ZIP contents from overwhelming the standalone browser app',
    ],
  },
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
