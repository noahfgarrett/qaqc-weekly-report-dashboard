import type { UpdateInfo } from '@/types'
import { isNewer } from '@/utils/semver'

interface GitHubAsset {
  name: string
  url: string
  browser_download_url: string
}

interface GitHubRelease {
  tag_name: string
  body?: string
  assets?: GitHubAsset[]
}

const TIMEOUT_MS = 5000

export async function checkForUpdate(): Promise<UpdateInfo | null> {
  try {
    const controller = new AbortController()
    const timer = window.setTimeout(() => controller.abort(), TIMEOUT_MS)
    const res = await fetch(`https://api.github.com/repos/${__GITHUB_REPO__}/releases/latest`, {
      signal: controller.signal,
      headers: { Accept: 'application/vnd.github.v3+json' },
    })
    window.clearTimeout(timer)
    if (!res.ok) return null

    const release = (await res.json()) as GitHubRelease
    const remoteVersion = release.tag_name.replace(/^v/, '')
    if (!isNewer(remoteVersion, __APP_VERSION__)) return null

    const htmlAsset = release.assets?.find((asset) => asset.name.toLowerCase().endsWith('.html'))
    if (!htmlAsset) return null

    return {
      version: remoteVersion,
      releaseNotes: release.body ?? '',
      downloadUrl: htmlAsset.browser_download_url,
      htmlDownloadUrl: htmlAsset.browser_download_url,
      assetApiUrl: htmlAsset.url,
      assetName: htmlAsset.name,
    }
  } catch {
    return null
  }
}

export async function downloadUpdate(info: UpdateInfo): Promise<void> {
  const a = document.createElement('a')
  a.href = info.htmlDownloadUrl ?? info.downloadUrl
  a.download = info.assetName
  a.target = '_blank'
  a.rel = 'noopener noreferrer'
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
}
