import { toCanvas } from 'html-to-image'

function waitForPaint(): Promise<void> {
  return new Promise((resolve) => {
    requestAnimationFrame(() => requestAnimationFrame(() => resolve()))
  })
}

async function renderElementPng(element: HTMLElement): Promise<Blob> {
  await document.fonts.ready
  await waitForPaint()
  const bounds = element.getBoundingClientRect()
  // Copy computed styles so the browser paints shadows and charts at their live size.
  const canvas = await toCanvas(element, {
    backgroundColor: '#ffffff',
    pixelRatio: 3,
    width: bounds.width,
    height: bounds.height,
    style: {
      position: 'relative',
      top: '0',
      left: '0',
      right: 'auto',
      bottom: 'auto',
      margin: '0',
      transform: 'none',
      outline: 'none',
      transition: 'none',
    },
    filter: (candidate) => !candidate.classList?.contains('chart-tooltip'),
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
