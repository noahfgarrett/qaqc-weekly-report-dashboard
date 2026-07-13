import pptxgen from 'pptxgenjs'
import type { AgingBucket, ElectricalPoint, MonthlyIssuePoint, ReportModel, WeeklyIssuePoint, WeldingPoint } from '@/types'
import { compactNumber, deltaLabel, percent } from '@/utils/format'

const SLIDE_W = 13.333
const SLIDE_H = 7.5
const SAFE_TOP = 0.58
const SAFE_BOTTOM = 0.54
const SAFE_H = SLIDE_H - SAFE_TOP - SAFE_BOTTOM
// Mirrors DESIGN.md tokens. Series keys keep their old names but new semantic
// values so existing draw calls remap: cyan=Opened(red), mint=Closed(green),
// amber=Remaining, teal=brand series (slate-blue), coral=alert.
const C = {
  ink: '1B2530',
  graphite: '28323F',
  text: '1B2530',
  muted: '6A7583',
  faint: 'DCE3EC',
  hairline: 'EBEEF3',
  panel: 'F2F5F9',
  surface: 'F7F9FC',
  white: 'FFFFFF',
  accent: '2E5AAC',
  accentStrong: '244A93',
  accentSoft: 'EAF0FA',
  track: 'E4E9F0',
  teal: '2E5AAC',
  cyan: 'EC6152',
  mint: '0D6331',
  amber: 'C2870B',
  coral: 'D03B3B',
  obsidian: '1B2530',
}

function addHeader(slide: pptxgen.Slide, title: string, report: ReportModel): void {
  slide.background = { color: C.white }
  // Slate-blue brand mark, matching the app header
  slide.addShape('roundRect', {
    x: 0.55,
    y: SAFE_TOP,
    w: 0.44,
    h: 0.44,
    rectRadius: 0.08,
    fill: { color: C.accent },
    line: { color: C.accent },
  })
  slide.addText(title, {
    x: 1.12,
    y: SAFE_TOP + 0.02,
    w: 6.0,
    h: 0.4,
    fontFace: 'Aptos Display',
    fontSize: 18,
    bold: true,
    color: C.text,
    valign: 'middle',
    margin: 0,
  })
  const pills = [
    `OAC Through ${report.reportWeek.label}`,
    report.source === 'files' ? 'Imported Files' : 'Preview Layout',
    `Report Week ${report.reportWeek.label}`,
  ]
  pills.forEach((pill, index) => {
    slide.addShape('roundRect', {
      x: 7.05 + index * 1.78,
      y: SAFE_TOP + 0.09,
      w: 1.62,
      h: 0.28,
      rectRadius: 0.06,
      fill: { color: C.panel },
      line: { color: C.faint },
    })
    slide.addText(pill, {
      x: 7.11 + index * 1.78,
      y: SAFE_TOP + 0.145,
      w: 1.5,
      h: 0.14,
      fontSize: 6.8,
      color: C.muted,
      bold: true,
      align: 'center',
      margin: 0,
      fit: 'shrink',
    })
  })
  // Hairline under the header
  addLine(slide, 0.55, SAFE_TOP + 0.54, 12.78, SAFE_TOP + 0.54, C.faint, 1)
}

function addFooter(slide: pptxgen.Slide, report: ReportModel): void {
  slide.addText(`Generated ${report.generatedAt.toLocaleDateString()} | v${__APP_VERSION__}`, {
    x: 0.55,
    y: SAFE_TOP + SAFE_H - 0.16,
    w: 5,
    h: 0.12,
    fontSize: 6.2,
    color: C.muted,
    margin: 0,
  })
}

function addKpiCard(slide: pptxgen.Slide, x: number, y: number, w: number, label: string, value: string, delta: string, color = C.accent): void {
  slide.addShape('roundRect', {
    x,
    y,
    w,
    h: 0.76,
    rectRadius: 0.1,
    fill: { color: C.surface },
    line: { color: C.faint },
    shadow: { type: 'outer', color: 'CBD5E1', opacity: 0.13, blur: 1, angle: 45, offset: 1 },
  })
  slide.addText(label, {
    x: x + 0.14,
    y: y + 0.11,
    w: w - 0.24,
    h: 0.12,
    fontSize: 6.6,
    bold: true,
    color: C.muted,
    margin: 0,
    fit: 'shrink',
  })
  slide.addText(value, {
    x: x + 0.14,
    y: y + 0.28,
    w: w - 0.24,
    h: 0.23,
    fontFace: 'Aptos Display',
    fontSize: 17,
    bold: true,
    color: C.text,
    margin: 0,
    fit: 'shrink',
  })
  slide.addText(delta, {
    x: x + 0.14,
    y: y + 0.56,
    w: w - 0.24,
    h: 0.1,
    fontSize: 6.4,
    bold: true,
    color,
    margin: 0,
    fit: 'shrink',
  })
}

function maxOf(values: number[]): number {
  return Math.max(1, ...values)
}

function addPanel(slide: pptxgen.Slide, x: number, y: number, w: number, h: number, title: string): void {
  slide.addShape('roundRect', {
    x,
    y,
    w,
    h,
    rectRadius: 0.12,
    fill: { color: C.white },
    line: { color: C.faint },
    shadow: { type: 'outer', color: 'CBD5E1', opacity: 0.18, blur: 1, angle: 45, offset: 1 },
  })
  slide.addText(title, {
    x: x + 0.18,
    y: y + 0.12,
    w: w - 0.36,
    h: 0.16,
    fontSize: 9,
    bold: true,
    color: C.text,
    margin: 0,
  })
}

function addLine(slide: pptxgen.Slide, x1: number, y1: number, x2: number, y2: number, color: string, width = 1, dash?: 'dash'): void {
  slide.addShape('line', {
    x: x1,
    y: y1,
    w: x2 - x1,
    h: y2 - y1,
    line: { color, width, dashType: dash },
  })
}

function addIssueTrend(slide: pptxgen.Slide, data: WeeklyIssuePoint[], x: number, y: number, w: number, h: number): void {
  addPanel(slide, x, y, w, h, 'Issues by Work Week')
  const visible = data.slice(-30)
  const chartX = x + 0.42
  const chartW = w - 0.72
  const groupW = chartW / Math.max(visible.length, 1)
  const maxBar = maxOf(visible.flatMap((d) => [d.opened, d.closed]))
  const maxLine = maxOf(visible.map((d) => d.remainingOpen))
  // Two stacked plots sharing the x-axis (no dual y-axis), matching the app.
  const lineTop = y + 0.5
  const lineH = h - 1.98
  const barTop = y + h - 1.12
  const barH = h - 2.28

  slide.addText('Remaining Open   Opened   Closed', {
    x: x + 0.42,
    y: y + 0.31,
    w: 3,
    h: 0.11,
    fontSize: 6,
    color: C.muted,
    margin: 0,
  })

  // Top plot: Remaining Open backlog line
  addLine(slide, chartX, lineTop + lineH, chartX + chartW, lineTop + lineH, 'D8DEE7', 0.7)
  visible.forEach((point, index) => {
    if (index === 0) return
    const prev = visible[index - 1]
    addLine(
      slide,
      chartX + (index - 0.5) * groupW,
      lineTop + lineH - (prev.remainingOpen / maxLine) * lineH,
      chartX + (index + 0.5) * groupW,
      lineTop + lineH - (point.remainingOpen / maxLine) * lineH,
      C.amber,
      1.4,
    )
  })
  const last = visible[visible.length - 1]
  if (last) {
    slide.addText(`Open ${last.remainingOpen}`, {
      x: chartX + chartW - 0.95,
      y: Math.max(lineTop, lineTop + lineH - (last.remainingOpen / maxLine) * lineH - 0.17),
      w: 0.9,
      h: 0.12,
      fontSize: 6.5,
      bold: true,
      color: C.amber,
      align: 'right',
      margin: 0,
    })
  }

  // Bottom plot: weekly Opened / Closed bars
  addLine(slide, chartX, barTop + barH, chartX + chartW, barTop + barH, 'D8DEE7', 0.7)
  visible.forEach((point, index) => {
    const gx = chartX + index * groupW + groupW * 0.2
    const openedH = (point.opened / maxBar) * barH
    const closedH = (point.closed / maxBar) * barH
    slide.addShape('roundRect', { x: gx, y: barTop + barH - openedH, w: groupW * 0.24, h: openedH, rectRadius: 0.02, fill: { color: C.cyan }, line: { color: C.cyan } })
    slide.addShape('roundRect', { x: gx + groupW * 0.3, y: barTop + barH - closedH, w: groupW * 0.24, h: closedH, rectRadius: 0.02, fill: { color: C.mint }, line: { color: C.mint } })
    if (index % 5 === 0 || index === visible.length - 1) {
      slide.addText(point.workWeek.replace('WW', ''), {
        x: chartX + index * groupW - 0.06,
        y: barTop + barH + 0.05,
        w: 0.44,
        h: 0.1,
        fontSize: 5,
        color: C.muted,
        align: 'center',
        margin: 0,
      })
    }
  })
}

function addMonthly(slide: pptxgen.Slide, data: MonthlyIssuePoint[], x: number, y: number, w: number, h: number): void {
  addPanel(slide, x, y, w, h, 'Cumulative Opened vs Closed')
  const chartX = x + 0.35
  const chartY = y + 0.54
  const chartW = w - 0.62
  const chartH = h - 0.85
  const visible = data.slice(-9)
  const maxY = maxOf(visible.flatMap((d) => [d.opened, d.closed]))
  const step = chartW / Math.max(1, visible.length - 1)
  addLine(slide, chartX, chartY + chartH, chartX + chartW, chartY + chartH, 'D8DEE7', 0.7)
  visible.forEach((point, index) => {
    if (index === 0) return
    const prev = visible[index - 1]
    const x1 = chartX + (index - 1) * step
    const x2 = chartX + index * step
    addLine(slide, x1, chartY + chartH - (prev.opened / maxY) * chartH, x2, chartY + chartH - (point.opened / maxY) * chartH, C.cyan, 1.5)
    addLine(slide, x1, chartY + chartH - (prev.closed / maxY) * chartH, x2, chartY + chartH - (point.closed / maxY) * chartH, C.mint, 1.5)
  })
  visible.forEach((point, index) => {
    if (index % 2 === 0 || index === visible.length - 1) {
      slide.addText(point.month, {
        x: chartX + index * step - 0.18,
        y: chartY + chartH + 0.06,
        w: 0.52,
        h: 0.1,
        fontSize: 5.5,
        color: C.muted,
        margin: 0,
      })
    }
  })
}

function addAging(slide: pptxgen.Slide, data: AgingBucket[], x: number, y: number, w: number, h: number): void {
  addPanel(slide, x, y, w, h, 'Issue Aging')
  const maxY = maxOf(data.map((d) => d.count))
  data.forEach((bucket, index) => {
    const yy = y + 0.55 + index * 0.42
    slide.addText(bucket.label, { x: x + 0.26, y: yy, w: 1.1, h: 0.12, fontSize: 6.5, color: C.muted, margin: 0 })
    slide.addShape('roundRect', {
      x: x + 1.25,
      y: yy - 0.02,
      w: (bucket.count / maxY) * (w - 1.82),
      h: 0.18,
      rectRadius: 0.04,
      fill: { color: bucket.color.replace('#', '') },
      line: { color: bucket.color.replace('#', '') },
    })
    slide.addText(compactNumber(bucket.count), { x: x + w - 0.44, y: yy, w: 0.25, h: 0.1, fontSize: 6.5, bold: true, color: C.text, margin: 0 })
  })
}

function addSimpleLine(slide: pptxgen.Slide, data: ElectricalPoint[], x: number, y: number, w: number, h: number): void {
  addPanel(slide, x, y, w, h, 'Electrical Inspections by Work Week')
  slide.addText('Final Inspections   Issue Found', { x: x + 0.45, y: y + 0.32, w: 2.8, h: 0.11, fontSize: 6, color: C.muted, margin: 0 })
  const chartX = x + 0.45
  const chartY = y + 0.58
  const chartW = w - 0.78
  const chartH = h - 1.05
  const visible = data.slice(-24)
  const maxY = maxOf(visible.map((d) => d.finals))
  const step = chartW / Math.max(1, visible.length - 1)
  addLine(slide, chartX, chartY + chartH, chartX + chartW, chartY + chartH, 'D8DEE7', 0.7)
  visible.forEach((point, index) => {
    if (index > 0) {
      const prev = visible[index - 1]
      addLine(
        slide,
        chartX + (index - 1) * step,
        chartY + chartH - (prev.finals / maxY) * chartH,
        chartX + index * step,
        chartY + chartH - (point.finals / maxY) * chartH,
        C.teal,
        1.5,
      )
    }
    if (point.issuesFound > 0) {
      slide.addShape('ellipse', {
        x: chartX + index * step - 0.04,
        y: chartY + chartH - (point.finals / maxY) * chartH - 0.04,
        w: 0.08,
        h: 0.08,
        fill: { color: C.coral },
        line: { color: C.coral },
      })
    }
    if (index % 4 === 0 || index === visible.length - 1) {
      slide.addText(point.workWeek.replace('WW', ''), { x: chartX + index * step - 0.04, y: chartY + chartH + 0.08, w: 0.35, h: 0.1, fontSize: 5, color: C.muted, rotate: 315, margin: 0 })
    }
  })
  slide.addText('Inspection Count', { x: x + 0.15, y: y + 1.1, w: 0.15, h: 1.2, rotate: 270, fontSize: 5.5, color: C.muted, margin: 0 })
  slide.addText('Work Week Observed', { x: x + 2.2, y: y + h - 0.18, w: 1.4, h: 0.1, fontSize: 5.5, color: C.muted, margin: 0 })
}

function addWelding(slide: pptxgen.Slide, data: WeldingPoint[], x: number, y: number, w: number, h: number): void {
  addPanel(slide, x, y, w, h, 'Welding Signoffs by Work Week')
  const chartX = x + 0.42
  const chartY = y + 0.58
  const chartW = w - 0.75
  const chartH = h - 1.05
  const visible = data.slice(-24)
  // Single axis (weld count): signed welds sit inside the total-weld track.
  const maxY = maxOf(visible.map((d) => d.total))
  const groupW = chartW / visible.length
  const barW = Math.min(0.17, groupW * 0.56)
  addLine(slide, chartX, chartY + chartH, chartX + chartW, chartY + chartH, 'D8DEE7', 0.7)
  visible.forEach((point, index) => {
    const bx = chartX + index * groupW + (groupW - barW) / 2
    const totalH = (point.total / maxY) * (chartH * 0.9)
    const signedH = (point.signed / maxY) * (chartH * 0.9)
    slide.addShape('roundRect', { x: bx, y: chartY + chartH - totalH, w: barW, h: totalH, rectRadius: 0.02, fill: { color: C.track }, line: { color: C.track } })
    slide.addShape('roundRect', { x: bx, y: chartY + chartH - signedH, w: barW, h: signedH, rectRadius: 0.02, fill: { color: C.mint }, line: { color: C.mint } })
    if (point.issuesCreated > 0) {
      slide.addShape('ellipse', { x: bx + barW + 0.01, y: chartY + chartH - totalH - 0.02, w: 0.07, h: 0.07, fill: { color: C.coral }, line: { color: C.coral } })
    }
    if (index === visible.length - 1) {
      slide.addText(percent(point.signoffRate, 0), { x: bx - 0.12, y: chartY + chartH - totalH - 0.17, w: barW + 0.36, h: 0.12, fontSize: 6, bold: true, color: C.accentStrong, align: 'center', margin: 0 })
    }
    if (index % 4 === 0 || index === visible.length - 1) {
      slide.addText(point.workWeek.replace('WW', ''), { x: chartX + index * groupW - 0.05, y: chartY + chartH + 0.08, w: 0.44, h: 0.1, fontSize: 5, color: C.muted, align: 'center', margin: 0 })
    }
  })
  slide.addText('Signed   Total welds   Issue created', { x: x + 0.4, y: y + 0.32, w: 3.2, h: 0.11, fontSize: 6, color: C.muted, margin: 0 })
  slide.addText('Weld Count', { x: x + 0.15, y: y + 1.18, w: 0.15, h: 1, rotate: 270, fontSize: 5.5, color: C.muted, margin: 0 })
}

function addIssueTable(slide: pptxgen.Slide, report: ReportModel): void {
  const x = 0.55
  const y = SAFE_TOP + 1.02
  const columns = [
    ['ID', 0.72],
    ['Subtype', 1.1],
    ['Status', 0.85],
    ['Title', 3.2],
    ['Contractor', 1.45],
    ['Work Week', 0.88],
    ['Created On', 0.82],
    ['Work Week Closed', 1.04],
    ['Due Date', 0.72],
  ] as const
  let cursor = x
  slide.addShape('roundRect', { x, y: y - 0.05, w: 12.25, h: 4.9, rectRadius: 0.1, fill: { color: C.white }, line: { color: C.faint } })
  slide.addShape('roundRect', { x, y, w: 12.25, h: 0.3, rectRadius: 0.06, fill: { color: C.panel }, line: { color: C.faint } })
  columns.forEach(([label, width]) => {
    slide.addText(label.toUpperCase(), { x: cursor + 0.05, y: y + 0.04, w: width - 0.1, h: 0.12, fontSize: 5.6, bold: true, color: C.muted, charSpacing: 0.4, margin: 0, fit: 'shrink' })
    cursor += width
  })
  const rows = report.issueTable.slice(0, 16)
  rows.forEach((row, rowIndex) => {
    const yy = y + 0.35 + rowIndex * 0.27
    slide.addShape('rect', {
      x,
      y: yy - 0.03,
      w: 12.25,
      h: 0.25,
      fill: { color: rowIndex % 2 === 0 ? 'F8FAFC' : 'FFFFFF' },
      line: { color: 'F1F5F9', transparency: 50 },
    })
    const values = [row.id, row.subtype, row.status, row.title, row.contractor, row.workWeek, row.createdOn, row.workWeekClosed, row.dueDate]
    cursor = x
    columns.forEach(([, width], index) => {
      const statusColor = row.status.toLowerCase().includes('closed') ? C.mint : row.status.toLowerCase().includes('overdue') ? C.coral : C.teal
      slide.addText(values[index], {
        x: cursor + 0.05,
        y: yy + 0.03,
        w: width - 0.1,
        h: 0.1,
        fontSize: 5.7,
        bold: index === 0 || index === 2,
        color: index === 2 ? statusColor : C.text,
        margin: 0,
        fit: 'shrink',
      })
      cursor += width
    })
  })
}

export async function exportReportDeck(report: ReportModel): Promise<void> {
  const pptx = new pptxgen()
  pptx.layout = 'LAYOUT_WIDE'
  pptx.author = 'QA/QC Weekly Report Dashboard'
  pptx.subject = `Weekly report through ${report.reportWeek.label}`
  pptx.title = `QA/QC Weekly Report ${report.reportWeek.label}`
  pptx.company = 'Generated Dashboard'
  pptx.theme = {
    headFontFace: 'Aptos Display',
    bodyFontFace: 'Aptos',
  }

  const slide1 = pptx.addSlide()
  addHeader(slide1, 'Weekly QA/QC Report', report)
  const kpiY = SAFE_TOP + 0.82
  report.kpis.forEach((metric, index) => {
    addKpiCard(slide1, 0.55 + (index % 4) * 3.08, kpiY + Math.floor(index / 4) * 0.86, 2.82, metric.label, metric.value, metric.deltaLabel, metric.tone === 'bad' ? C.coral : metric.tone === 'warn' ? C.amber : metric.tone === 'good' ? C.mint : C.teal)
  })
  addIssueTrend(slide1, report.issueTrend, 0.55, SAFE_TOP + 2.72, 7.6, 3.05)
  addMonthly(slide1, report.monthlyTrend, 8.38, SAFE_TOP + 2.72, 4.0, 1.45)
  addAging(slide1, report.aging, 8.38, SAFE_TOP + 4.36, 4.0, 1.41)
  addFooter(slide1, report)

  const slide2 = pptx.addSlide()
  addHeader(slide2, 'BIM Issues Detail', report)
  addKpiCard(slide2, 0.55, SAFE_TOP + 0.82, 2.7, 'Open Carryover', compactNumber(report.issueTable.filter((row) => row.group === 'Open Carryover').length), '', C.teal)
  addKpiCard(slide2, 3.47, SAFE_TOP + 0.82, 2.7, 'Closed in Report Week', compactNumber(report.issueTable.filter((row) => row.group === 'Closed in Report Week' || row.group === 'Opened + Closed in Report Week').length), '', C.mint)
  addKpiCard(slide2, 6.39, SAFE_TOP + 0.82, 2.7, 'Opened + Closed', compactNumber(report.issueTable.filter((row) => row.group === 'Opened + Closed in Report Week').length), '', C.cyan)
  addKpiCard(slide2, 9.31, SAFE_TOP + 0.82, 2.7, 'Closed This Week', compactNumber(report.issueTable.filter((row) => row.group === 'Closed This Week').length), '', C.amber)
  addIssueTable(slide2, report)
  addFooter(slide2, report)

  const slide3 = pptx.addSlide()
  addHeader(slide3, 'Inspections & Welding Signoffs', report)
  const chips = [
    ['Electrical Finals', compactNumber(report.summary.electricalFinals), deltaLabel(report.summary.deltas.electricalFinals)],
    ['Electrical Issues Found', compactNumber(report.summary.electricalIssuesFound), deltaLabel(report.summary.deltas.electricalIssuesFound)],
    ['Welds Checked', compactNumber(report.summary.weldsChecked), deltaLabel(report.summary.deltas.weldsChecked)],
    ['Welds Signed', compactNumber(report.summary.weldsSigned), deltaLabel(report.summary.deltas.weldsSigned)],
    ['Avg Sign-off %', percent(report.summary.avgSignoffRate, 1), deltaLabel(report.summary.deltas.avgSignoffRate)],
  ]
  chips.forEach(([label, value, delta], index) => addKpiCard(slide3, 0.55 + index * 2.48, SAFE_TOP + 0.82, 2.22, label, value, delta, index === 1 ? C.amber : C.teal))
  addSimpleLine(slide3, report.electrical, 0.55, SAFE_TOP + 2.08, 5.95, 3.74)
  addWelding(slide3, report.welding, 6.82, SAFE_TOP + 2.08, 5.95, 3.74)
  addFooter(slide3, report)

  await pptx.writeFile({ fileName: `QAQC Weekly Report ${report.reportWeek.label.replace("'", '-')}.pptx` })
}
