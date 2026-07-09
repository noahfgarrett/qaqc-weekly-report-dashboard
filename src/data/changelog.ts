export interface ChangelogEntry {
  version: string
  date: string
  type: 'feature' | 'fix' | 'major'
  notes: string[]
}

export const CHANGELOG: ChangelogEntry[] = [
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
