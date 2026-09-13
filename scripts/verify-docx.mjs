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

import JSZip from 'jszip'

import { normalizeBlocks, parseMd, resolveSpec, toBase64, toMd } from '../dist-lib/wordtohtml.mjs'

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
  // 下划线的两种形态都要走一遍：整段带下划线（Word 会报 Underline=single），
  // 以及与加粗混排的一段（Word 报 wdUndefined，用来确认读回值确实区分得出来）
  '__本段整段带下划线__',
  '',
  '**加粗**与__下划线__并存的一段。',
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
  '% 附件一：资产核查明细表',
  '',
  // 分页符紧挨着分节符（故意留的连排）：分页符后面没有段落可挂，必须退回成
  // 独立段落里的 w:br w:type="page"，否则换页会被静默丢掉。Word 每次验 P1
  // 都会真的打开一次这种形状。
  '===',
  '---',
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

const buffer = Buffer.from(
  await toBase64(model, spec, { title: '爱康光电资产核查情况说明' }),
  'base64',
)

const outPath = join(root, '.qwen', 'tmp', 'verify.docx')
mkdirSync(dirname(outPath), { recursive: true })
writeFileSync(outPath, buffer)

/*
 * 行单位段距必须在 OOXML 层面就能看见。
 *
 * Word 的 styles[].SpaceBefore 报的是 w:before/w:after 这对后备值，看的是「多少磅」，
 * 证明不了这一段在 Word 里被当成「行」—— 判定「行」与「磅」的是 w:beforeLines /
 * w:afterLines 在不在。所以这里直接看写进文件的字节。
 */
{
  const zip = await JSZip.loadAsync(buffer)
  const stylesXml = await zip.file('word/styles.xml').async('string')
  const docXml = await zip.file('word/document.xml').async('string')

  const problems = []

  // 正文（docDefaults）也要写成行单位
  const defaults = /<w:pPrDefault\b[\s\S]*?<\/w:pPrDefault>/.exec(stylesXml)?.[0] ?? ''
  if (!/w:beforeLines=/.test(defaults)) problems.push('docDefaults（正文）没写 beforeLines')

  // 规格表里每条样式（不含正文）都该带上
  const styleIds = Object.entries(spec.styles)
    .filter(([key]) => key !== 'body')
    .map(([, s]) => s.id)
  const bare = styleIds.filter((id) => {
    const block = new RegExp(
      `<w:style\\b[^>]*?w:styleId="${id}"[^>]*>[\\s\\S]*?</w:style>`,
    ).exec(stylesXml)?.[0]
    return !block || !/w:beforeLines=/.test(block)
  })
  if (bare.length > 0) problems.push(`这些样式没写 beforeLines：${bare.join('、')}`)

  const grids = [...docXml.matchAll(/<w:docGrid\b[^>]*>/g)].map((m) => m[0])
  if (grids.length === 0) problems.push('document.xml 里没有 w:docGrid（「行」就没有基准）')

  /*
   * 下划线必须在 OOXML 层面看得见。Word 报的 Font.Underline 只能说明「渲染成了下划线」，
   * 说不出是哪一层写出来的（样式里也可能有），所以这里直接看写进文件的字节：
   * 每一处带下划线的 inline 都应该有且只有一个 <w:u w:val="single"/>。
   */
  const underlineTags = [...docXml.matchAll(/<w:u\b[^>]*\/>/g)].map((m) => m[0])
  const modelUnderlined = model.blocks
    .filter((b) => b.t === 'textBlock')
    .flatMap((b) => b.inlines)
    .filter((i) => i.t === 'text' && i.underline).length
  // 只有内置样本才要求「必须验到下划线」：这个脚本也接受外部源码（verify:pages 把
  // 界面里的源码原样喂进来），而别人的源码里没有下划线是正常的。
  if (sourceIndex < 0 && modelUnderlined === 0) {
    problems.push('内置样本里没有带下划线的文字 —— 这一项等于没验')
  }
  if (underlineTags.length !== modelUnderlined) {
    problems.push(
      `下划线 run 数不符：document.xml 里 ${underlineTags.length} 处，模型里 ${modelUnderlined} 处`,
    )
  }
  const notSingle = underlineTags.filter((tag) => !/w:val="single"/.test(tag))
  if (notSingle.length > 0) problems.push(`这些 w:u 不是 single：${notSingle.join('、')}`)

  if (problems.length > 0) {
    console.error('[FAIL] docx 的字节与模型/规格表对不上：')
    for (const p of problems) console.error(`  - ${p}`)
    process.exit(1)
  }
  console.log(
    `[ok] 行单位段距：${styleIds.length + 1} 处 spacing 带 beforeLines/afterLines，` +
      `${grids.length} 处 docGrid（${grids[0]}）`,
  )
  console.log(`[ok] 下划线：${underlineTags.length} 处 w:u，全部 val="single"`)
}

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
