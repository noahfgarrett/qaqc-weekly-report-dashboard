import type { UpdateInfo } from '../types'

interface SaveUpdateFile {
  (blob: Blob, filename: string): Promise<void> | void
}

interface UpdateDownloadDeps {
  fetchImpl?: typeof fetch
  saveImpl?: SaveUpdateFile
  decompressImpl?: (blob: Blob) => Promise<Blob>
  shouldUseRawDownload?: () => boolean
}

export interface UpdateDownloadResult {
  downloadedAssetName: string
  savedFilename: string
  usedCompressedAsset: boolean
  usedRawDownloadFallback?: boolean
}

type DecompressionStreamCtor = new (format: 'gzip') => TransformStream<Uint8Array, Uint8Array>

function getHtmlFilename(assetName: string, fallbackAssetName?: string): string {
  if (fallbackAssetName) return fallbackAssetName
  return assetName.toLowerCase().endsWith('.gz') ? assetName.slice(0, -3) : assetName
}

function getDecompressionStream(): DecompressionStreamCtor | undefined {
  return (globalThis as { DecompressionStream?: DecompressionStreamCtor }).DecompressionStream
}

function isStandaloneFileContext(): boolean {
  return typeof window !== 'undefined' && (window.location.protocol === 'file:' || window.location.origin === 'null')
}

async function useRawDownload(
  info: UpdateInfo,
  fetchImpl: typeof fetch,
  saveImpl: SaveUpdateFile,
): Promise<UpdateDownloadResult> {
  const filename = info.downloadKind === 'gzip-html' ? info.fallbackAssetName : info.assetName
  if (!filename) throw new Error('No packaged HTML filename is attached to this release.')
  const htmlBlob = await fetchAssetBlob(info.rawDownloadUrl, fetchImpl, 'text/html,application/octet-stream')
  await saveImpl(new Blob([htmlBlob], { type: 'text/html' }), filename)
  return {
    downloadedAssetName: filename,
    savedFilename: filename,
    usedCompressedAsset: false,
    usedRawDownloadFallback: true,
  }
}

function saveBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = filename
  document.body.appendChild(anchor)
  anchor.click()
  document.body.removeChild(anchor)
  window.setTimeout(() => URL.revokeObjectURL(url), 60_000)
}

async function fetchAssetBlob(url: string, fetchImpl: typeof fetch, accept: string): Promise<Blob> {
  const response = await fetchImpl(url, { headers: { Accept: accept } })
  if (!response.ok) throw new Error(`Update download failed: HTTP ${String(response.status)}`)
  return response.blob()
}

async function fetchHtmlAsset(
  assetApiUrl: string,
  downloadUrl: string,
  fetchImpl: typeof fetch,
): Promise<Blob> {
  try {
    return await fetchAssetBlob(assetApiUrl, fetchImpl, 'application/octet-stream')
  } catch {
    return fetchAssetBlob(downloadUrl, fetchImpl, 'text/html,application/octet-stream')
  }
}

export function canDecompressGzipInBrowser(): boolean {
  return typeof getDecompressionStream() === 'function'
}

export async function decompressGzipHtml(blob: Blob): Promise<Blob> {
  const DecompressionStream = getDecompressionStream()
  if (!DecompressionStream) throw new Error('This browser cannot decompress gzip updates.')
  const decompressed = blob.stream().pipeThrough(new DecompressionStream('gzip'))
  const htmlBlob = await new Response(decompressed).blob()
  return new Blob([htmlBlob], { type: 'text/html' })
}

export async function downloadUpdateFile(
  info: UpdateInfo,
  deps: UpdateDownloadDeps = {},
): Promise<UpdateDownloadResult> {
  const fetchImpl = deps.fetchImpl ?? fetch
  const saveImpl = deps.saveImpl ?? saveBlob
  const decompressImpl = deps.decompressImpl ?? decompressGzipHtml
  const shouldUseRawDownload = deps.shouldUseRawDownload ?? isStandaloneFileContext

  if (shouldUseRawDownload()) return useRawDownload(info, fetchImpl, saveImpl)

  if (info.downloadKind === 'gzip-html' && canDecompressGzipInBrowser()) {
    try {
      const compressedBlob = await fetchAssetBlob(info.assetApiUrl, fetchImpl, 'application/octet-stream')
      const htmlBlob = await decompressImpl(compressedBlob)
      const filename = getHtmlFilename(info.assetName, info.fallbackAssetName)
      await saveImpl(htmlBlob, filename)
      return {
        downloadedAssetName: info.assetName,
        savedFilename: filename,
        usedCompressedAsset: true,
      }
    } catch {
      // The plain HTML fallback below remains entirely in-app via fetch + Blob.
    }
  }

  const htmlApiUrl = info.downloadKind === 'gzip-html' ? info.fallbackAssetApiUrl : info.assetApiUrl
  const htmlDownloadUrl = info.downloadKind === 'gzip-html' ? info.fallbackDownloadUrl : info.downloadUrl
  const htmlAssetName = info.downloadKind === 'gzip-html' ? info.fallbackAssetName : info.assetName
  if (!htmlApiUrl || !htmlDownloadUrl || !htmlAssetName) {
    throw new Error('No compatible packaged HTML update is attached to this release.')
  }

  try {
    const htmlBlob = await fetchHtmlAsset(htmlApiUrl, htmlDownloadUrl, fetchImpl)
    await saveImpl(new Blob([htmlBlob], { type: 'text/html' }), htmlAssetName)
    return {
      downloadedAssetName: htmlAssetName,
      savedFilename: htmlAssetName,
      usedCompressedAsset: false,
    }
  } catch {
    return useRawDownload(info, fetchImpl, saveImpl)
  }
}
