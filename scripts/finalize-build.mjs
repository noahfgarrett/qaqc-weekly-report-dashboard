import { copyFileSync, existsSync, mkdirSync, readFileSync, rmSync } from 'node:fs'
import { resolve } from 'node:path'

const root = process.cwd()
const { version } = JSON.parse(readFileSync(resolve(root, 'package.json'), 'utf8'))
const packagedName = `QAQC-Weekly-Report-Dashboard-v${version}.html`
const distIndex = resolve(root, 'dist/index.html')
const distNamed = resolve(root, 'dist', packagedName)
const outputDir = resolve(root, 'outputs')
const outputFile = resolve(outputDir, packagedName)
const legacyOutputFile = resolve(outputDir, 'QAQC-Weekly-Report-Dashboard.html')
const releaseDir = resolve(root, 'release')
const releaseFile = resolve(releaseDir, 'QAQC-Weekly-Report-Dashboard.html')

if (!existsSync(distIndex)) {
  throw new Error('dist/index.html was not generated.')
}

mkdirSync(outputDir, { recursive: true })
mkdirSync(releaseDir, { recursive: true })
copyFileSync(distIndex, distNamed)
copyFileSync(distIndex, outputFile)
copyFileSync(distIndex, releaseFile)
rmSync(legacyOutputFile, { force: true })

console.log(`Wrote ${distNamed}`)
console.log(`Wrote ${outputFile}`)
console.log(`Wrote ${releaseFile}`)
