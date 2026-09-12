/**
 * P1 验收第 1 步：把一份覆盖全部语法特性的样本编译成 docx。
 *
 * 顺带做「模型 → md → 模型」的往返一致性检查，确认解析与序列化没有暗坑。
 * 产物写两份：
 *   - .qwen/tmp/verify.docx  留档，你也能直接打开看
 *   - 系统临时目录的 wtp-verify.docx  纯 ASCII 路径，给 Word COM 脚本用
 *     （避免中文路径在 cmd.exe → powershell.exe 之间被改写编码）
 */

import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

import { Packer } from 'docx'

import { buildDocument, normalizeBlocks, parseMd, resolveSpec, toMd } from '../dist-lib/wordtohtml.mjs'

const here = dirname(fileURLToPath(import.meta.url))
const root = resolve(here, '..')

const SAMPLE = [
  '# 关于爱康光电资产核查情况的说明',
  '',
  '@ 苏州市公安局经济犯罪侦查支队：',
  '',
  '我方于2026年9月1日收到你单位《调取证据通知书》，现就有关事项说明如下。',
  '',
  '## 资产核查情况',
  '',
  '经核查，{+我单位}对债务人名下资产进行了全面梳理，**重点**如下：',
  '',
  '### 不动产',
  '',
  '#### 房产',
  '',
  '- 位于苏州市工业园区的厂房一处',
  '- 位于吴中区的办公用房一处',
  '',
  '! 上述资产均已办理抵押登记',
  '',
  '{-该笔债务已经清偿}',
  '',
  '相关日期以{红|通知书}记载为准[[通知书原件|日期需与通知书原件核对]]。',
  '',
  '---',
  '',
  '## 附件说明',
  '',
  '本节为附件，页码重新起算。',
  '',
  '>> 江苏爱康光电破产管理人',
  '>> 2026年9月12日',
  '',
].join('\n')

// 允许用外部源码覆盖内置样本：verify-docx.mjs --source <文件>。
// demo 页的页数对账就靠它把界面里的源码原样喂进来。
const sourceIndex = process.argv.indexOf('--source')
const source =
  sourceIndex >= 0 && process.argv[sourceIndex + 1]
    ? readFileSync(process.argv[sourceIndex + 1], 'utf8')
    : SAMPLE

const FIXED_DATE = new Date('2026-09-12T08:00:00Z')
const now = () => FIXED_DATE

const spec = resolveSpec()
const model = parseMd(source, { author: '张三', now })

const md2 = toMd(model)
const model2 = parseMd(md2, { author: '张三', now })

const before = JSON.stringify(normalizeBlocks(model))
const after = JSON.stringify(normalizeBlocks(model2))
if (before !== after) {
  console.error('[FAIL] 往返不一致：模型 → md → 模型 之后结构发生了变化')
  console.error('  原始:', before)
  console.error('  往返:', after)
  process.exit(1)
}
console.log('[ok] 往返一致性通过（模型 → md → 模型）')

const doc = buildDocument(model, spec, { title: '爱康光电资产核查情况说明' })
const buffer = await Packer.toBuffer(doc)

const outPath = join(root, '.qwen', 'tmp', 'verify.docx')
mkdirSync(dirname(outPath), { recursive: true })
writeFileSync(outPath, buffer)

// 临时副本名带上 pid：Word 退出后会短暂占住文件，用固定名会让下一次运行
// 撞上 EBUSY。每次换一个新名字就永远不会冲突，跑完由 verify-p1.mjs 清理。
const asciiPath = join(tmpdir(), `wtp-verify-${process.pid}.docx`)
writeFileSync(asciiPath, buffer)

const modelPath = join(root, '.qwen', 'tmp', 'model.json')
writeFileSync(modelPath, JSON.stringify({ spec, model }, null, 2), 'utf8')

console.log(
  JSON.stringify(
    {
      ok: true,
      bytes: buffer.length,
      outPath,
      asciiPath,
      modelPath,
      blocks: model.blocks.length,
      comments: model.comments.length,
      md: md2,
    },
    null,
    2,
  ),
)
