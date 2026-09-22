import html2canvas from 'html2canvas'

function waitForPaint(): Promise<void> {
  return new Promise((resolve) => {
    requestAnimationFrame(() => requestAnimationFrame(() => resolve()))
  })
}

async function renderElementPng(element: HTMLElement): Promise<Blob> {
  await document.fonts.ready
  await waitForPaint()
  const canvas = await html2canvas(element, {
    backgroundColor: '#ffffff',
    scale: 3,
    useCORS: true,
    logging: false,
    windowWidth: element.scrollWidth,
    windowHeight: element.scrollHeight,
    ignoreElements: (candidate) => candidate.classList?.contains('chart-tooltip') ?? false,
  })
  try {
    const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, 'image/png'))
    if (!blob) throw new Error('The report image could not be encoded.')
    return blob
  } finally {
    canvas.width = 1
    canvas.height = 1
  }
}

export async function copyElementAsPng(element: HTMLElement): Promise<void> {
  if (!navigator.clipboard?.write || typeof ClipboardItem === 'undefined') {
    throw new Error('Image clipboard access is not available in this browser.')
  }
  const image = renderElementPng(element)
  await navigator.clipboard.write([new ClipboardItem({ 'image/png': image })])
}
