import html2canvas from 'html2canvas'
import { jsPDF } from 'jspdf'
import type { ReportModel } from '@/types'

export async function exportSlidesPdf(report: ReportModel): Promise<void> {
  const slides = Array.from(document.querySelectorAll<HTMLElement>('[data-export-slide]'))
  if (slides.length === 0) throw new Error('No slides are available for PDF export.')

  const pdf = new jsPDF({
    orientation: 'landscape',
    unit: 'in',
    format: [13.333, 7.5],
    compress: true,
  })

  for (const [index, slide] of slides.entries()) {
    if (index > 0) pdf.addPage([13.333, 7.5], 'landscape')
    const canvas = await html2canvas(slide, {
      backgroundColor: '#ffffff',
      scale: 2,
      useCORS: true,
      logging: false,
      windowWidth: slide.scrollWidth,
      windowHeight: slide.scrollHeight,
    })
    const img = canvas.toDataURL('image/png')
    pdf.addImage(img, 'PNG', 0, 0, 13.333, 7.5, undefined, 'FAST')
  }

  pdf.save(`QAQC Weekly Report ${report.reportWeek.label.replace("'", '-')}.pdf`)
}
