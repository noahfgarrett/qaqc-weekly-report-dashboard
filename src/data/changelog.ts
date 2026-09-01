export interface ChangelogEntry {
  version: string
  date: string
  type: 'feature' | 'fix' | 'major'
  notes: string[]
}

export const CHANGELOG: ChangelogEntry[] = [
  {
    version: '1.2.1',
    date: '2026-09-01T19:06:24Z',
    type: 'fix',
    notes: [
      'The missing-data review is now non-blocking, so any subset of corrections can be applied while unresolved blanks are skipped',
      'Blank Contractor and General Contractor suggestions remain editable and can be accepted independently',
      'Mechanical and Electrical inspection rows no longer flag Subtype because it is not used by their report calculations',
      'Welding rows no longer flag Contractor, Discipline, or Subtype because those fields do not drive welding metrics',
      'Review wording now clearly distinguishes optional blanks from corrections ready to apply',
    ],
  },
  {
    version: '1.2.0',
    date: '2026-09-01T14:31:53Z',
    type: 'feature',
    notes: [
      'Added a pre-report data quality review that flags blank fields used by dashboard calculations, filters, and Issue Detail',
      'Blank Contractor and General Contractor cells are suggested as Bechtel while remaining fully editable before applying corrections',
      'The review identifies the source sheet, record, row, and field; unresolved required values must be completed before the report opens',
      'Intentional blanks such as unsigned welds, empty issue indicators, and closure dates on open issues remain valid and are not flagged',
      'Corrections are applied to the in-memory imported sheets so the dashboard, PowerPoint, and PDF all use the reviewed values',
    ],
  },
  {
    version: '1.1.6',
    date: '2026-08-26T13:38:00Z',
    type: 'fix',
    notes: [
      'Justine now always opens with All contractors instead of inheriting the Bechtel selection from OAC',
      'The previous OAC contractor selection is remembered and restored when switching back to OAC',
      'Saved Justine sessions discard stale OAC contractor restrictions while retaining every available contractor in the slicer',
    ],
  },
  {
    version: '1.1.5',
    date: '2026-08-26T13:24:00Z',
    type: 'fix',
    notes: [
      'Restored BIM Issues Detail to the same full 16:9 canvas height used by Overview and Inspections & Welding',
      'The live Issue Detail table now uses twelve complete rows at the original readable sizing, without vertically compressing the slide',
      'Reporting-week activity remains first and only unused rows are filled with the most recent open issues',
    ],
  },
  {
    version: '1.1.4',
    date: '2026-08-26T13:11:00Z',
    type: 'fix',
    notes: [
      'Restored the original Issue Detail slide, card, title, and table sizing so report content remains comfortably readable and unclipped',
      'The live canvas preserves the original readable typography and card sizing',
      'The newest open issues fill only genuinely unused rows after reporting-week issues, with no typography compression to accommodate additional backlog',
      'A full report page receives no supplemental open issues; a page with four empty rows receives only the four most recent open issues',
    ],
  },
  {
    version: '1.1.3',
    date: '2026-08-26T12:58:10Z',
    type: 'feature',
    notes: [
      'BIM Issues Detail now uses fixed 14-row in-app pages so the entire report canvas fits on screen for clean snips without table scrolling',
      'Compact previous and next controls make additional Issue Detail pages easy to review while keeping every row fully visible',
      'Replaced report-count text with a clear Reporting Week and Most Recent Open Issues row-color legend across the app and exports',
    ],
  },
  {
    version: '1.1.2',
    date: '2026-08-26T12:36:31Z',
    type: 'feature',
    notes: [
      'BIM Issues Detail now fills unused report-page rows with the newest still-open or Pending issues after the report-week activity rows',
      'Supplemental open issues use a distinct slate row treatment so they are available for discussion without changing report-week metrics',
      'HTML, PDF, and PowerPoint use only the highest-ID supplemental issues that fit without creating extra backlog-only pages',
    ],
  },
  {
    version: '1.1.1',
    date: '2026-08-26T12:08:36Z',
    type: 'feature',
    notes: [
      'Replaced Cumulative Opened vs Closed with an Open Issue Aging chart grouped into 0-1, 1-2, 2-3, 3-4, and 4+ week buckets',
      'Renamed the existing aging panel to Overall Issue Aging (Opened & Closed) for clearer scope',
      'Swapped the reporting-period and Project to Date KPI groups across the dashboard, PDF, and PowerPoint layouts',
      'The reporting-period KPI group now uses the lighter surface treatment, with Project to Date using the contrasting gray treatment',
      'Rebalanced the aging-chart row and introduced one slate-blue scale that darkens as issue age increases',
    ],
  },
  {
    version: '1.1.0',
    date: '2026-08-20T11:48:01Z',
    type: 'major',
    notes: [
      'Manual Issues Update now carries historical BIM ownership, Created On, and derived closed dates into the enriched ACC copy',
      'Historical Updated On becomes the driving closed date only when the matching BIM issue was Closed; open issues retain ACC closure timing',
      'Matched historical dates take priority, while unmatched ACC issues use Created On, Closed at, and Updated on fallbacks',
      'Issue #1021 uses Bechtel when its Contractor is blank, while any future populated Contractor remains authoritative',
      'The latest raw ACC export was reconciled row by row against BIM Issues Detail with no missing, unexpected, or misclassified report activity',
    ],
  },
  {
    version: '1.0.31',
    date: '2026-08-13T18:01:57Z',
    type: 'feature',
    notes: [
      'Added a persistent Justine reporting mode beside OAC and Manual in the weekly report toolbar',
      'Justine combines the previous and current work weeks across issue, inspection, SOR, electrical, and welding summaries',
      'Dashboard, BIM Issues Detail, PDF, and PowerPoint exports now share the selected reporting period and comparison window',
    ],
  },
  {
    version: '1.0.30',
    date: '2026-08-11T13:35:46Z',
    type: 'fix',
    notes: [
      'Closure Rate week-over-week comparisons now use green when the percentage improves and red when it declines',
      'An unchanged Closure Rate remains neutral across the dashboard and exported Project to Date cards',
      'Regression coverage verifies improving, declining, and unchanged Closure Rate scenarios',
    ],
  },
  {
    version: '1.0.29',
    date: '2026-08-11T13:05:31Z',
    type: 'feature',
    notes: [
      'Manual Issues Update is now a focused ACC enrichment workflow that matches IDs against a current reference log',
      'Only blank ACC Contractor and Discipline cells are filled from matching reference IDs; populated ACC values are never overwritten',
      'The download is a new ACC workbook copy with every row, column, worksheet, metadata value, and original row order preserved',
      'The enrichment summary now reports matched IDs, Contractor fills, Discipline fills, unmatched rows, and duplicate reference IDs',
    ],
  },
  {
    version: '1.0.28',
    date: '2026-08-11T11:32:42Z',
    type: 'fix',
    notes: [
      'ACC Type is now the primary source for BIM issue Subtype, with older Subtype headers retained as a fallback',
      'ACC rows with blank BIM360 legacy metadata now use Created By or Issue Owner for LotusWorks ownership while older curated logs remain compatible',
      'Excel serial dates and numeric date strings from General-formatted XLSX or CSV exports now resolve to the correct dates and work weeks',
      'Added a comprehensive ACC metadata matrix covering ownership, status aliases, legacy dates, Type priority, metrics, and BIM Issues Detail groups',
    ],
  },
  {
    version: '1.0.27',
    date: '2026-08-10T15:05:41Z',
    type: 'fix',
    notes: [
      'Added first-class support for BIM360_Created By, BIM360_Created On, and BIM360_Closed On legacy export columns',
      'Legacy LotusWorks ownership and historical dates now take priority, with the existing ACC method retained when all three legacy cells are blank',
      'Closed legacy issues without BIM360_Closed On use Updated on as their closure date, while unclosed issues remain open',
    ],
  },
  {
    version: '1.0.26',
    date: '2026-08-06T12:50:09Z',
    type: 'fix',
    notes: [
      'BIM Issues Detail is now limited to issues represented by its four activity cards',
      'The detail list includes issues opened in the reporting week, closed in the reporting week, opened and closed in the reporting week, or closed in the current week',
      'Older open issues are no longer included solely because their Updated on date falls within the reporting week',
    ],
  },
  {
    version: '1.0.25',
    date: '2026-08-06T12:28:06Z',
    type: 'fix',
    notes: [
      'BIM Issues Detail now includes open and Pending issues updated during the reporting week, even when they were created earlier',
      'Pending issues included through reporting-week activity continue to display as Open and count toward open-issue metrics',
      'Regression coverage verifies issue 1018-style activity across WW30 creation and WW31 update dates',
    ],
  },
  {
    version: '1.0.24',
    date: '2026-08-06T11:59:10Z',
    type: 'fix',
    notes: [
      'Pending BIM issues now display as Open in BIM Issues Detail, including PPTX and PDF exports',
      'Pending issues continue to count toward Total Issues Opened and Issues Remaining Open across all report metrics',
      'The Pending status remains available in the Status slicer for precise source-data filtering',
    ],
  },
  {
    version: '1.0.23',
    date: '2026-08-06T11:47:02Z',
    type: 'fix',
    notes: [
      'Manual Issues Update now preserves the original BIM Issues Log Updated on value for every existing ID',
      'Historical issue closure work weeks are no longer replaced by the ACC migration date',
      'Regression coverage verifies both Created on and Updated on remain in their original historical work weeks',
    ],
  },
  {
    version: '1.0.22',
    date: '2026-08-06T11:41:21Z',
    type: 'fix',
    notes: [
      'Manual Issues Update now preserves the original BIM Issues Log Created on value for every existing ID instead of accepting ACC migration dates',
      'ACC continues to update all other supported fields for existing issues, while genuinely new LotusWorks IDs receive their Created on value from ACC',
      'Regression coverage reproduces the July 24 Autodesk transfer-date problem and verifies that the original historical date still resolves to its correct work week',
    ],
  },
  {
    version: '1.0.21',
    date: '2026-08-06T11:30:53Z',
    type: 'fix',
    notes: [
      'Manual Issues Update now writes Created on, Updated on, and Due date as true Excel short-date cells instead of General-formatted values',
      'Updated BIM workbooks are sorted by numeric-aware descending ID so the highest BIM issue number appears first',
      'The descending sort moves complete worksheet rows together, preserving every related field while placing blank-ID rows last',
      'Workbook regression coverage now verifies date formats and BIM-1000-series ordering in the generated XLSX file',
    ],
  },
  {
    version: '1.0.20',
    date: '2026-08-06T11:23:30Z',
    type: 'fix',
    notes: [
      'Manual Issues Update now maps the ACC Type column into the BIM Issues Log Subtype column',
      'The ACC Category column is intentionally ignored while the current BIM Issues Log continues to use its existing Subtype header',
      'The workbook regression fixture now verifies the ACC-specific Type-to-Subtype translation for both updated and newly added issues',
    ],
  },
  {
    version: '1.0.19',
    date: '2026-08-06T10:53:08Z',
    type: 'feature',
    notes: [
      'Added a dedicated Manual Issues Update workspace for reconciling a current BIM Issues Log against the replacement ACC export',
      'Every ID already present in the BIM Issues Log is refreshed from ACC even when Autodesk transfer ownership replaced the original creator, while genuinely new IDs are added only when Created By contains LotusWorks',
      'The merge updates only the nine dashboard fields, preserves existing values when ACC sends a blank cell, and prevents duplicate or missing IDs from entering the log',
      'A review table and automatic Update Summary modal show updated, added, unchanged, and excluded records before the versioned workbook downloads locally',
    ],
  },
  {
    version: '1.0.18',
    date: '2026-08-05T18:06:48Z',
    type: 'fix',
    notes: [
      'BIM issues whose Status is Pending now explicitly count as open throughout totals, remaining-open backlog, weekly and monthly trends, aging, filtering, and detail exports',
      'Pending appears as its own Status slicer value and uses a consistent amber treatment in the HTML dashboard, PDF output, and PowerPoint issue table',
      'Added a focused regression check that verifies Pending contributes to opened and remaining totals but never to closed totals',
    ],
  },
  {
    version: '1.0.17',
    date: '2026-07-22T16:06:50Z',
    type: 'fix',
    notes: [
      'Electrical issue counts now include only rows whose Inspection Phase is Final, preventing non-final inspections from adding issue markers or summary totals',
    ],
  },
  {
    version: '1.0.16',
    date: '2026-07-15T17:00:00Z',
    type: 'feature',
    notes: [
      'The dashboard now checks GitHub Releases on launch, every 15 minutes, and when the app becomes visible again, automatically presenting newer packaged versions',
      'Update downloads are fetched from the release tag as CORS-safe packaged HTML and saved locally without navigating away from the dashboard or opening a GitHub page',
      'The Multitool-style update modal now includes dedicated Update and Changelog tabs, release notes, download progress and success states, and clear network error recovery',
      'The in-app changelog now uses expandable release entries, version types, dates, latest-version context, and a manual Check now control',
      'The persistent Updates button now displays an unobtrusive version badge whenever a newer packaged release is available',
    ],
  },
  {
    version: '1.0.15',
    date: '2026-07-15T15:00:00Z',
    type: 'fix',
    notes: [
      'With OAC enabled, BIM Issues Detail now includes only issues opened during the reporting week, closed during the reporting week, opened and closed within that reporting week, or closed during the current week',
      'The focused OAC activity window is shared by the HTML table, PDF export, and native PowerPoint export while manual slicer mode retains its broader issue-detail view',
      'The first OAC detail summary card now identifies issues opened during the reporting week instead of describing the removed historical carryover population',
      'High-resolution PDF capture now isolates and validates each page, automatically retries incomplete raster output, and prevents overlapping PDF and PowerPoint generation',
    ],
  },
  {
    version: '1.0.14',
    date: '2026-07-15T13:50:00Z',
    type: 'fix',
    notes: [
      'The Welding Signoffs chart now labels only the active reporting-week sign-off percentage while retaining the complete historical line and hover details',
      'The same focused reporting-week label is used in the native PowerPoint export for a cleaner field-report chart',
    ],
  },
  {
    version: '1.0.13',
    date: '2026-07-15T13:25:31Z',
    type: 'feature',
    notes: [
      'The report-week KPI panel now keeps its values and sparklines clean by omitting previous-week comparison labels',
      'Redundant Issues / Week and Remaining Open captions were removed from the Issues by Work Week plot while retaining both numeric scales and the legend',
      'Open-issue comparisons now use red for increases, green for decreases, and neutral styling when unchanged',
      'BIM issue exports now paginate at complete 14-row boundaries so an issue never splits across PDF or PowerPoint slides',
      'BIM detail summary cards use clearer descriptions, including Issues Remaining Open and explicit report-week/current-week closure labels',
      'Electrical and Welding charts now show rolling 30-week windows with compact WW## axes and current work-week context in the slide header',
      'Welding now uses Total Welds terminology, collision-aware data labels, and a front-layer dotted red 10% baseline with an inline axis label',
      'Detail cards, field cards, chart panels, and the BIM issue table now use consistent restrained drop shadows',
    ],
  },
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
