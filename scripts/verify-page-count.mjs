/**
 * 预览页数 vs Word 页数 —— 整条链路最有价值的一致性检查。
 *
 * 分页器、量测、样式的单测都只能证明「内部自洽」。真正要回答的是：
 * 我们的 A4 预览和 Word 打开同一个 docx 之后的页数是否一致。
 * 这需要两个引擎（Chromium 的排版与 Word 的排版）在同样的行距、段距、
 * 孤行规则下得出同一个分页结果。
 *
 * 前置：先跑 verify-browser.mjs（它产出 demo-source.md 与 pagination-report.json）。
 * 用法：npm run verify:pages
 */

import { spawnSync } from 'node:child_process'
import { existsSync, readFileSync, rmSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const root = resolve(here, '..')
const tmpDir = join(root, '.qwen', 'tmp')

function run(label, command, args, options = {}) {
  const result = spawnSync(command, args, {
    cwd: root,
    encoding: 'utf8',
    shell: false,
    ...options,
  })
  if (result.status !== 0) {
    console.error(`[FAIL] ${label} 退出码 ${result.status}`)
    if (result.stderr) console.error(result.stderr)
    process.exit(result.status ?? 1)
  }
  return result
}

const browserPath = join(tmpDir, 'pagination-report.json')
const sourcePath = join(tmpDir, 'demo-source.md')
if (!existsSync(browserPath) || !existsSync(sourcePath)) {
  console.error('缺少前置产物，请先运行：npm run verify:p2')
  process.exit(2)
}

const browser = JSON.parse(readFileSync(browserPath, 'utf8').replace(/^\uFEFF/, ''))

console.log('### 用同一份源码生成 docx')
const generated = run('verify-docx', process.execPath, [
  'scripts/verify-docx.mjs',
  '--source',
  sourcePath,
])
const info = JSON.parse(generated.stdout.slice(generated.stdout.indexOf('{')))

const dumpPath = info.asciiPath.replace(/\.docx$/, '.json')

try {
  console.log('### 让 Word 报告页数')
  run('check-docx', 'powershell.exe', [
    '-NoProfile',
    '-ExecutionPolicy',
    'Bypass',
    '-File',
    join(root, 'scripts', 'check-docx.ps1'),
    '-Path',
    info.asciiPath,
    '-OutPath',
    dumpPath,
  ])

  const dump = JSON.parse(readFileSync(dumpPath, 'utf8').replace(/^\uFEFF/, ''))

  console.log('\n=== 页数对账 ===')
  console.log(`  浏览器预览：${browser.pageCount} 页，页码序列 ${browser.pageNumbers.join(',')}`)
  console.log(`  Word：${dump.pageCount} 页，分节 ${dump.sectionCount} 节`)

  const failures = []
  if (dump.pageCount !== browser.pageCount) {
    failures.push(`页数不一致：预览 ${browser.pageCount}，Word ${dump.pageCount}`)
  }
  if (dump.sectionCount !== browser.pageNumbers.filter((n) => n === '1').length) {
    failures.push(
      `分节数不一致：Word ${dump.sectionCount} 节，预览里出现 ${browser.pageNumbers.filter((n) => n === '1').length} 次页码 1`,
    )
  }

  if (failures.length > 0) {
    console.error('\n[FAIL]')
    for (const f of failures) console.error(`  - ${f}`)
    process.exit(1)
  }
  console.log('\n[PASS] 预览分页与 Word 的分页结果一致（页数、分节重编号均吻合）。')
} finally {
  for (const path of [info.asciiPath, dumpPath]) {
    if (!existsSync(path)) continue
    try {
      rmSync(path, { force: true })
    } catch {
      // Word 可能还占着，留着无妨
    }
  }
}
