/**
 * P1 的单一验收入口：一条命令跑完整条链路。
 *
 *   1. node scripts/verify-docx.mjs   生成 docx + 往返一致性检查
 *   2. powershell check-docx.ps1      让 Word 自己报出实际生效的样式与结构
 *   3. node scripts/assert-docx.mjs   把 Word 的说法与规格表逐项对账
 *
 * 之所以要有这个编排脚本：Word 打开过的文件会被短暂占住，临时路径必须每次
 * 唯一（由 verify-docx.mjs 带 pid 生成），跑完还要清理 —— 这些散在 shell 里
 * 迟早会出错，集中到一处更好维护。
 *
 * 用法：npm run verify:p1
 */

import { spawnSync } from 'node:child_process'
import { existsSync, rmSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { tmpdir } from 'node:os'

const here = dirname(fileURLToPath(import.meta.url))
const root = resolve(here, '..')

function run(label, command, args, options = {}) {
  const result = spawnSync(command, args, {
    cwd: root,
    stdio: options.capture ? 'pipe' : 'inherit',
    encoding: 'utf8',
    shell: false,
    ...options,
  })
  if (result.status !== 0) {
    console.error(`\n[FAIL] ${label} 退出码 ${result.status}`)
    if (options.capture && result.stderr) console.error(result.stderr)
    process.exit(result.status ?? 1)
  }
  return result
}

console.log('### 1/3 生成 docx 并做往返一致性检查')
const generated = run('verify-docx', process.execPath, ['scripts/verify-docx.mjs'], {
  capture: true,
})
const info = JSON.parse(generated.stdout.slice(generated.stdout.indexOf('{')))
console.log(`  ok  ${info.asciiPath}（${info.bytes} 字节，${info.blocks} 块，${info.comments} 条批注）`)

const dumpPath = info.asciiPath.replace(/\.docx$/, '.json')
const cleanup = [info.asciiPath, dumpPath]

try {
  console.log('\n### 2/3 让 Word 报告它实际识别到的样式与结构')
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
  console.log('  ok  Word 已完成读取并退出')

  console.log('\n### 3/3 与规格表逐项对账')
  run('assert-docx', process.execPath, [
    'scripts/assert-docx.mjs',
    dumpPath,
    join(root, '.qwen', 'tmp', 'model.json'),
  ])
} finally {
  for (const path of cleanup) {
    if (!existsSync(path)) continue
    try {
      rmSync(path, { force: true })
    } catch {
      // Word 偶尔还要多占几秒，留着不影响结果
      console.log(`  （临时文件暂被占用，未删除：${path}）`)
    }
  }
}
