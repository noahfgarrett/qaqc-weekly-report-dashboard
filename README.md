# QA/QC Weekly Report Dashboard

Premium HTML dashboard for Smartsheet-backed weekly QA/QC reporting.

## What It Does

- Connects to Smartsheet with a saved local token.
- Auto-maps the default sheets:
  - BIM Issues Log
  - Mechanical / Process Inspection Log
  - Electrical Inspection Log
  - Welding Signoffs by Work Week
- Defaults OAC reporting to all work weeks through the previous completed work week.
- Provides slicers for Work Week, Discipline, Contractor, Subtype, and Status.
- Exports a GC-safe three-slide report deck as PPTX or PDF.
- Checks release metadata for packaged HTML app updates and downloads the new HTML file directly.

## Local Development

```bash
npm install
npm run dev
```

## Build

```bash
npm run build
```

The single-file HTML output is written to:

```text
outputs/QAQC-Weekly-Report-Dashboard.html
```

## Release Workflow

1. Update `package.json` version and `src/data/changelog.ts`.
2. Run `npm run build`.
3. Create a GitHub release tag.
4. Attach `outputs/QAQC-Weekly-Report-Dashboard.html` to the release.

The in-app updater expects the latest release to include an `.html` asset. The modal downloads that asset directly as a browser download; users do not need to open GitHub.

## Security

Smartsheet tokens are stored only in the browser's local storage on the user's device. Tokens are not committed, bundled, or uploaded by this app.
