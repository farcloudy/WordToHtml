/**
 * 预览页数 vs Word 页数 —— 整条链路最有价值的一致性检查。
 *
 * 分页器、量测、样式的单测都只能证明「内部自洽」。真正要回答的是：
 * 我们的 A4 预览和 Word 打开同一个 docx 之后的页数是否一致。
 * 这需要两个引擎（Chromium 的排版与 Word 的排版）在同样的行距、段距、
 * 孤行规则下得出同一个分页结果。
 *
 * 两套文件模板都要过：公文标准的版心更矮，正是最容易和 Word 数字对不上的情形。
 *   1/2 用 demo 页导出的源码，按每套模板各生成一份 docx，与预览页数对齐；
 *   2/2 默认模板之外的模板再用内置样本过一次 Word 逐项对账（assert-docx）——
 *        默认模板那套在 verify:p1 里已经跑过，这里不重复启动 Word。
 *
 * 前置：先跑 verify-browser.mjs（它产出 demo-source.md 与每套模板的分页报告）。
 * 用法：npm run verify:pages
 */

import { spawnSync } from 'node:child_process'
import { existsSync, readFileSync, rmSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

import { DOC_TEMPLATES } from '../dist-lib/wordtohtml.mjs'

const here = dirname(fileURLToPath(import.meta.url))
const root = resolve(here, '..')
const tmpDir = join(root, '.qwen', 'tmp')
const modelPath = join(tmpDir, 'model.json')

const failures = []

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

function rmQuiet(path) {
  if (!existsSync(path)) return
  try {
    rmSync(path, { force: true })
  } catch {
    // Word 可能还占着，留着无妨（文件名带 pid 与时间戳，不会撞车）
  }
}

/**
 * 生成 docx → 让 Word 读回。临时文件由调用方删（下一步可能还要拿 dump 去对账）。
 * 返回 Word 报出来的结构（dump）与生成信息。
 */
function wordRoundTrip(template, sourceArgs) {
  const generated = run(`verify-docx（${template.key}）`, process.execPath, [
    'scripts/verify-docx.mjs',
    ...sourceArgs,
    '--template',
    template.key,
  ])
  const info = JSON.parse(generated.stdout.slice(generated.stdout.indexOf('{')))
  const dumpPath = info.asciiPath.replace(/\.docx$/, '.json')
  run(`check-docx（${template.key}）`, 'powershell.exe', [
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
  return { info, dump, dumpPath }
}

function dropRoundTrip(round) {
  rmQuiet(round.info.asciiPath)
  rmQuiet(round.dumpPath)
}

const browserPath = join(tmpDir, 'pagination-report.json')
const sourcePath = join(tmpDir, 'demo-source.md')
if (!existsSync(browserPath) || !existsSync(sourcePath)) {
  console.error('缺少前置产物，请先运行：npm run verify:p2')
  process.exit(2)
}

const browser = JSON.parse(readFileSync(browserPath, 'utf8').replace(/^\uFEFF/, ''))
const reports = browser.templates ?? {}

console.log('### 1/2 用同一份源码逐套模板生成 docx，比对页数与分节')
for (const template of DOC_TEMPLATES) {
  const preview = reports[template.key]
  if (!preview) {
    failures.push(`分页报告里没有模板 ${template.key}（verify-browser 是否跑的是旧版本？）`)
    continue
  }

  const round = wordRoundTrip(template, ['--source', sourcePath])
  const { dump } = round
  dropRoundTrip(round)

  /*
   * 页码 1 的出现次数不能拿来当节数（W5 起新插的分节符默认「关联前节 + 不重排」）。
   * 期望值从**模型**推：页码 1 恰好出现「首节 1 次 + 声明了从 1 开始的节数」次；
   * 分节数则直接比 Word 的 sectionCount 与模型的节数。
   * 两个数都由 verify-browser 用 resolveSections() 算好写进报告（与预览/导出同源）。
   */
  const restarts = preview.pageNumbers.filter((n) => n === '1').length
  const restartCount = preview.restartCount ?? 0
  const expectedOnes = 1 + restartCount
  const expectedSections = preview.sectionCount ?? restarts

  console.log(`\n=== 页数对账「${template.label}」===`)
  console.log(`  浏览器预览：${preview.pageCount} 页，页码序列 ${preview.pageNumbers.join(',')}`)
  console.log(`  Word：${dump.pageCount} 页，分节 ${dump.sectionCount} 节`)

  if (dump.pageCount !== preview.pageCount) {
    failures.push(
      `「${template.label}」页数不一致：预览 ${preview.pageCount}，Word ${dump.pageCount}`,
    )
  }
  if (dump.sectionCount !== expectedSections) {
    failures.push(
      `「${template.label}」分节数不一致：Word ${dump.sectionCount} 节，模型里 ${expectedSections} 节`,
    )
  }
  if (restarts !== expectedOnes) {
    failures.push(
      `「${template.label}」页码 1 的出现次数与模型不符：预览里 ${restarts} 次，` +
        `模型推出应为 ${expectedOnes} 次（首节 1 次 + ${restartCount} 个声明了「从 1 开始」的节）`,
    )
  }
}

// 两套模板的页数必须不同 —— 一样就说明切换模板没有真的重量（页边距变了版心就变了）
if (DOC_TEMPLATES.length > 1) {
  const counts = DOC_TEMPLATES.map((t) => reports[t.key]?.pageCount ?? -1)
  const expected = DOC_TEMPLATES.map((t) => `${t.label}=${reports[t.key]?.pageCount}`).join('，')
  if (new Set(counts).size <= 1) {
    failures.push(`两套模板的预览页数相同（${expected}）—— 切换模板没有整篇重量`)
  } else {
    console.log(`\n  ok   两套模板的预览页数不同：${expected}`)
  }
}

/*
 * 默认模板之外的模板，再用内置样本过一次 Word 逐项对账：
 * 样式、段落结构、纸张与页边距、分节都会逐项比。
 * 内置样本才有修订与批注（assert-docx 按它写死了条数），所以这里不能用 demo 源码。
 */
for (const template of DOC_TEMPLATES.slice(1)) {
  console.log(`\n### 2/2 内置样本 + Word 逐项对账「${template.label}」`)
  const round = wordRoundTrip(template, [])
  try {
    // assert-docx 的逐项结论直接进日志：只报退出码看不出它到底对了哪些项
    const asserted = run(`assert-docx（${template.key}）`, process.execPath, [
      'scripts/assert-docx.mjs',
      round.dumpPath,
      modelPath,
      '--template',
      template.key,
    ])
    process.stdout.write(asserted.stdout ?? '')
  } finally {
    dropRoundTrip(round)
  }
}

console.log('')
if (failures.length > 0) {
  console.error('[FAIL]')
  for (const f of failures) console.error(`  - ${f}`)
  process.exit(1)
}
console.log('[PASS] 每套模板的预览分页与 Word 的分页结果一致（页数、节数、页码重排次数均吻合），')
console.log('       且两套模板的页数确实不同（切模板整篇重量的证据）。')
