/**
 * 直接读 docx 的 OOXML 内容。
 *
 * Word COM 报的是「Word 理解的语义」，这个脚本报的是「文件里实际写的字节」。
 * 两者都看，才能区分「我们没写对」和「Word 没按我们写的理解」。
 *
 * 用法：
 *   node scripts/dump-ooxml.mjs <docx>                      # 只列部件清单
 *   node scripts/dump-ooxml.mjs <docx> "pgNumType|w:start"   # 列出匹配片段
 */

import { readFileSync } from 'node:fs'

import JSZip from 'jszip'

const [, , docxPath, pattern] = process.argv
if (!docxPath) {
  console.error('用法: node scripts/dump-ooxml.mjs <docx> [正则]')
  process.exit(2)
}

const zip = await JSZip.loadAsync(readFileSync(docxPath))
const entries = Object.keys(zip.files)
  .filter((n) => !zip.files[n]?.dir)
  .sort()

if (!pattern) {
  console.log(`部件清单（${docxPath}）：`)
  for (const name of entries) {
    const raw = await zip.files[name]?.async('uint8array')
    console.log(`  ${name}  (${raw?.length ?? 0} bytes)`)
  }
  process.exit(0)
}

const re = new RegExp(pattern, 'g')
for (const name of entries) {
  const text = await zip.files[name]?.async('string')
  if (!text) continue
  const hits = [...text.matchAll(re)]
  if (hits.length === 0) continue
  console.log(`\n===== ${name}  命中 ${hits.length} 处 =====`)
  for (const hit of hits) {
    const at = hit.index ?? 0
    const from = Math.max(0, at - 120)
    const to = Math.min(text.length, at + 260)
    console.log(`  …${text.slice(from, to).replace(/\s+/g, ' ')}…`)
  }
}
