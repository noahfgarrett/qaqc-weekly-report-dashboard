import type { UpdateInfo } from '../types'
import { isNewer } from '../utils/semver'

export interface GitHubAsset {
  name: string
  url: string
  browser_download_url: string
}

interface GitHubRelease {
  tag_name: string
  body?: string
  assets?: GitHubAsset[]
}

export type SelectedUpdateAsset = Pick<
  UpdateInfo,
  | 'downloadUrl'
  | 'assetApiUrl'
  | 'assetName'
  | 'downloadKind'
  | 'fallbackDownloadUrl'
  | 'fallbackAssetApiUrl'
  | 'fallbackAssetName'
>

const TIMEOUT_MS = 7000

export function selectUpdateAsset(assets: GitHubAsset[] | undefined): SelectedUpdateAsset | null {
  const htmlAsset = assets?.find((asset) => asset.name.toLowerCase().endsWith('.html'))
  if (!htmlAsset) return null

  const gzipAsset = assets?.find((asset) => asset.name.toLowerCase().endsWith('.html.gz'))
  if (!gzipAsset) {
    return {
      downloadUrl: htmlAsset.browser_download_url,
      assetApiUrl: htmlAsset.url,
      assetName: htmlAsset.name,
      downloadKind: 'html',
    }
  }

  return {
    downloadUrl: gzipAsset.browser_download_url,
    assetApiUrl: gzipAsset.url,
    assetName: gzipAsset.name,
    downloadKind: 'gzip-html',
    fallbackDownloadUrl: htmlAsset.browser_download_url,
    fallbackAssetApiUrl: htmlAsset.url,
    fallbackAssetName: htmlAsset.name,
  }
}

export async function checkForUpdate(): Promise<UpdateInfo | null> {
  const controller = new AbortController()
  const timer = window.setTimeout(() => controller.abort(), TIMEOUT_MS)
  try {
    const response = await fetch(`https://api.github.com/repos/${__GITHUB_REPO__}/releases/latest`, {
      signal: controller.signal,
      headers: { Accept: 'application/vnd.github+json' },
    })
    if (!response.ok) return null

    const release = (await response.json()) as GitHubRelease
    const remoteVersion = release.tag_name.replace(/^v/, '')
    if (!isNewer(remoteVersion, __APP_VERSION__)) return null

    const selectedAsset = selectUpdateAsset(release.assets)
    if (!selectedAsset) return null

    return {
      version: remoteVersion,
      releaseNotes: release.body ?? '',
      rawDownloadUrl: `https://raw.githubusercontent.com/${__GITHUB_REPO__}/${encodeURIComponent(release.tag_name)}/release/QAQC-Weekly-Report-Dashboard.html`,
      ...selectedAsset,
    }
  } catch {
    return null
  } finally {
    window.clearTimeout(timer)
  }
}
