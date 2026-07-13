# PRODUCT.md — QA/QC Weekly Report Dashboard

## Register

**product** — design serves the task. This is a data tool, not a marketing surface.

## Product purpose

A standalone, single-file HTML dashboard that turns four weekly Smartsheet exports
(BIM Issues, Mechanical/Process, Electrical, Welding) into a reviewed QA/QC status
report for a construction project, and exports a General-Contractor-safe deck
(PPTX / PDF). All data stays in the browser session; nothing is uploaded.

## Users

- **Primary:** the project's QA/QC manager / engineer. Imports the weekly logs each
  Tuesday, reviews the numbers, then exports a report to send to the General
  Contractor. Fluent in construction terms (work weeks "WW28'2026", SORs,
  sign-off rate, finals). Wants to trust the numbers at a glance and hand off a
  clean deck without hand-editing.
- **Secondary (read-only):** the General Contractor and project leadership who
  receive the exported deck. They see the output, never the app.

## Scene (theme driver)

A QA/QC manager at a desk on a 24-27" monitor in an office under normal daytime
light, preparing a weekly report to send onward. Professional, focused, not a
glance-from-across-the-room wall display, not a 2am incident console.
→ **Light theme.** The app should also match the white printed/exported deck.

## Brand & tone

Serious, precise, quietly professional. An engineering instrument, not a
consumer app. Confidence through restraint and correct numbers, not decoration.
The output goes to an external contractor, so it must read as credible and
buttoned-up.

## Strategic principles

- The data is the only thing allowed to be loud. Chrome recedes.
- One disciplined brand accent (slate-blue). Red / amber / green mean *status*,
  never decoration.
- App and exported deck share one visual language — a GC seeing the deck and a
  manager seeing the app should see the same product.
- Correct-by-construction charts: colorblind-safe (validated), one y-axis each,
  legible at export resolution.

## Anti-references

- The current build's five-accent rainbow (teal + cyan + mint + amber + coral used
  together) and teal radial-gradient wash.
- Generic SaaS "hero metric + 8 identical cards + gradient accent" dashboards.
- Alarm-red everywhere: opened issues are normal operations, not a crisis.
