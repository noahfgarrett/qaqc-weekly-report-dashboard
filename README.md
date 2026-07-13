# QA/QC Weekly Report Dashboard

Standalone HTML dashboard for weekly QA/QC reporting from Smartsheet exports.

## What It Does

- Imports a ZIP, full folder, or individual `.xls`, `.xlsx`, and `.csv` files exported from Smartsheet.
- Auto-maps the four required logs by filename and column headers:
  - BIM Issues Log
  - Mechanical / Process Inspection Log
  - Electrical Inspection Log
  - Welding Signoffs by Work Week
- Defaults OAC reporting to all work weeks through the previous completed work week.
- Provides slicers for Work Week, Discipline, Contractor, Subtype, and Status.
- Exports a GC-safe three-slide report deck as PPTX or PDF.
- Checks release metadata for packaged HTML app updates and downloads the new HTML file directly.

## Weekly Workflow

1. Schedule the four Smartsheet logs to be emailed as Excel attachments each Tuesday.
2. Download the combined ZIP or attachments from the company email account.
3. Open `QAQC-Weekly-Report-Dashboard.html` and drop in the ZIP, its folder, or the four individual files.
4. Review the report and export the GC-safe PPTX or PDF.

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

The dashboard does not request or store a Smartsheet API token. Imported spreadsheet data stays in memory for the current browser session and is not persisted or uploaded. Legacy token data from earlier versions is removed automatically. The optional update checker only requests public GitHub release metadata.
