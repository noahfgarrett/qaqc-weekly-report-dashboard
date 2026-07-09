import { copyFileSync, existsSync, mkdirSync } from 'node:fs'
import { resolve } from 'node:path'

const root = process.cwd()
const distIndex = resolve(root, 'dist/index.html')
const distNamed = resolve(root, 'dist/QAQC-Weekly-Report-Dashboard.html')
const outputDir = resolve(root, 'outputs')
const outputFile = resolve(outputDir, 'QAQC-Weekly-Report-Dashboard.html')

if (!existsSync(distIndex)) {
  throw new Error('dist/index.html was not generated.')
}

mkdirSync(outputDir, { recursive: true })
copyFileSync(distIndex, distNamed)
copyFileSync(distIndex, outputFile)

console.log(`Wrote ${distNamed}`)
console.log(`Wrote ${outputFile}`)
