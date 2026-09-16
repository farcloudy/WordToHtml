/**
 * W8：库产物（dist-lib）能不能被使用方拿去用。
 *
 * build:lib 成功只说明「打包没报错」，不说明产物可用。这里在 node 里真 import 一次，
 * 断言三件事：
 *   1. 能 import —— 不抛异常（`.vue` 被编译进产物、依赖都 external 得对）；
 *   2. 导出了对外组件 `WtpEditor`（还有内部的 `WordPaper`）；
 *   3. 样式产物存在且非空 —— 组件的 `<style scoped>` 得跟着库一起发出去，
 *      否则使用方拿到的是一堆没有样式的 DOM。
 * 另按外部依赖再确认一条：`vue` 是 external（产物里 import 它，而不是打进一份自己的运行时）。
 *
 * 用法：npm run verify:lib（先 build:lib；package.json 里就是这么串的）
 */

import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const root = resolve(here, '..')
const libDir = join(root, 'dist-lib')
const entry = join(libDir, 'wordtohtml.mjs')

const failures = []

function ok(label, condition, detail = '') {
  if (condition) {
    console.log(`  ok   ${label}`)
    return true
  }
  failures.push(`${label}${detail ? ` — ${detail}` : ''}`)
  console.error(`  FAIL ${label}${detail ? ` — ${detail}` : ''}`)
  return false
}

console.log('\n=== 库产物（W8）===')

ok('库入口产物存在', existsSync(entry), entry)
if (!existsSync(entry)) {
  console.error('[FAIL] 先跑 npm run build:lib')
  process.exit(1)
}
const entrySource = readFileSync(entry, 'utf8')
ok('入口产物非空', statSync(entry).size > 0, `${statSync(entry).size} 字节`)

// 1/3：真 import 一次（用 file URL，避开 Windows 路径分隔符被当转义）
let mod = null
try {
  mod = await import(`file://${entry.replace(/\\/g, '/')}`)
  ok('node 里能 import 库产物（没抛异常）', true)
} catch (error) {
  ok('node 里能 import 库产物（没抛异常）', false, String(error))
}

if (mod) {
  // 2/3：导出的组件
  ok(
    '导出了对外组件 WtpEditor',
    mod.WtpEditor !== undefined && typeof mod.WtpEditor === 'object',
    typeof mod.WtpEditor,
  )
  ok(
    'WtpEditor 是被编译过的 SFC（__name 对得上）',
    mod.WtpEditor?.__name === 'WtpEditor',
    String(mod.WtpEditor?.__name),
  )
  ok(
    '内置的 WordPaper 也一并导出（高级用法）',
    mod.WordPaper !== undefined && typeof mod.WordPaper === 'object',
    typeof mod.WordPaper,
  )
  // 既有能力没被这次改动弄丢：抽查两个纯函数
  ok(
    '原有的纯函数导出仍在（parseMd / toMd）',
    typeof mod.parseMd === 'function' && typeof mod.toMd === 'function',
    `${typeof mod.parseMd}/${typeof mod.toMd}`,
  )
}

// 3/3：样式产物
const cssFiles = existsSync(libDir)
  ? readdirSync(libDir).filter((name) => name.toLowerCase().endsWith('.css'))
  : []
ok('产物里有 CSS 文件', cssFiles.length > 0, JSON.stringify(readdirSync(libDir)))
const cssName = cssFiles[0] ?? ''
const cssPath = cssName ? join(libDir, cssName) : ''
const cssText = cssPath && existsSync(cssPath) ? readFileSync(cssPath, 'utf8') : ''
ok('样式产物非空', cssText.length > 0, `${cssName} ${cssText.length} 字符`)
ok('样式里确实带上了组件的外壳规则（.bar / .canvas）', cssText.includes('.bar') && cssText.includes('.canvas'))
ok('样式里确实带上了纸张规则（.wtp-）', cssText.includes('.wtp-'))
console.log(`  --   CSS 产物：${cssName}（${cssText.length} 字符）`)

// 外部依赖：vue 必须留在外面，否则使用方会有两份运行时
ok('vue 是外部依赖（产物里 import "vue" 而不是内联）', /\bfrom\s*["']vue["']/.test(entrySource))
console.log(`  --   入口产物 ${statSync(entry).size} 字节`)

console.log('')
if (failures.length > 0) {
  console.error(`[FAIL] 共 ${failures.length} 项不符：`)
  for (const f of failures) console.error(`  - ${f}`)
  process.exit(1)
}
console.log(`[PASS] 库产物：可 import、导出 WtpEditor（含样式产物 ${cssName}）、vue 保持 external`)
