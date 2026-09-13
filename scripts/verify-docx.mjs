/**
 * P1 验收第 1 步：把一份覆盖全部语法特性的样本编译成 docx。
 *
 * 顺带做「模型 → md → 模型」的往返一致性检查，确认解析与序列化没有暗坑。
 * 产物写两份：
 *   - .qwen/tmp/verify.docx  留档，你也能直接打开看
 *   - 系统临时目录的 wtp-verify-<pid>-<时间戳>.docx  纯 ASCII 路径，给 Word COM 脚本用
 *     （避免中文路径在 cmd.exe → powershell.exe 之间被改写编码）
 *
 * 用法：node scripts/verify-docx.mjs [--source <文件>] [--template <key>]
 *   不传 --source 用内置样本；不传 --template 用 DOC_TEMPLATES 的第一套。
 */

import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

import JSZip from 'jszip'

import {
  DOC_TEMPLATES,
  allInlineHolders,
  lengthToPx,
  normalizeBlocks,
  parseMd,
  resolveSpec,
  toBase64,
  toMd,
} from '../dist-lib/wordtohtml.mjs'

const here = dirname(fileURLToPath(import.meta.url))
const root = resolve(here, '..')

/** 取一个 `--flag value` 形式的命令行参数 */
function argValue(flag) {
  const index = process.argv.indexOf(flag)
  return index >= 0 ? process.argv[index + 1] : undefined
}

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
  // 表格：覆盖全部形状 —— unit 行（右对齐、整行合并、无框）、加粗的列标题行、
  // 若干数据行、note 行（左对齐 + 顶端对齐、整行合并、无框），以及带 `|` 与 `\`
  // 的格（验转义往返）。表格里刻意不放修订与批注：assert-docx 的「修订 2 条 /
  // 批注 1 条」是写死的，没必要为了表格去扩大战线。
  ':::table minLines=2',
  '> 单位：元',
  '| **项目** | **金额** |',
  '| 甲资产 | 1,234.00 |',
  // 格内软换行（{br} → <w:br/>）：Word 读回来是垂直制表符 chr(11)，assert-docx 按这个对账
  '| 乙资产{br}（含附属设施） | 5,678.90 |',
  // 格内 `{红|…}` 的那枚竖线是内容、不是列分隔符：切格必须跳 `{}` 指令（只数反斜杠会把它切开）
  '| {红|丙资产} | 9,999.00 |',
  // `\|` 是内容里的竖线、`\\` 是内容里的反斜杠：切格必须数反斜杠，否则会被切开
  '| 备注\\|说明 | 含\\\\反斜杠 |',
  '< 注：以上金额不含税',
  ':::',
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
const source = (() => {
  const path = argValue('--source')
  return path ? readFileSync(path, 'utf8') : null
})()

/*
 * 用哪套文件模板：verify-docx.mjs [--template <key>]，默认第一套。
 * 页边距是模板的一部分，所以这里必须整份 spec 从 DOC_TEMPLATES 推导，
 * 不能只换页边距 —— 否则「简易公文格式」的 docx 就不是它自己了。
 */
const templateKey = argValue('--template')
const template = templateKey
  ? DOC_TEMPLATES.find((t) => t.key === templateKey)
  : DOC_TEMPLATES[0]
if (!template) {
  console.error(
    `[FAIL] 未知模板：${templateKey}（可选：${DOC_TEMPLATES.map((t) => t.key).join(' / ')}）`,
  )
  process.exit(2)
}
console.log(`[ok] 文件模板：${template.label}（${template.key}）`)

const usedBuiltinSample = source === null
const bodySource = source ?? SAMPLE

const FIXED_DATE = new Date('2026-09-12T08:00:00Z')
const now = () => FIXED_DATE

const spec = resolveSpec(template.spec)
const model = parseMd(bodySource, { author: '张三', now })

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
  if (usedBuiltinSample && modelUnderlined === 0) {
    problems.push('内置样本里没有带下划线的文字 —— 这一项等于没验')
  }
  if (underlineTags.length !== modelUnderlined) {
    problems.push(
      `下划线 run 数不符：document.xml 里 ${underlineTags.length} 处，模型里 ${modelUnderlined} 处`,
    )
  }
  const notSingle = underlineTags.filter((tag) => !/w:val="single"/.test(tag))
  if (notSingle.length > 0) problems.push(`这些 w:u 不是 single：${notSingle.join('、')}`)

  /*
   * 软换行必须在字节层看得见。Word 报的 Range.Text 里它是 chr(11)，那只说明「换行生效了」，
   * 看不出是哪一层写出来的（段落标记、单元格标记也都带控制符），所以直接数字节：
   * 模型里有几枚软换行，document.xml 里就该有几个**不带属性**的 <w:br/>。
   * 正则刻意不放宽：`<w:br w:type="page"/>` 是分页符，不能混进来。
   */
  const softBreakTags = [...docXml.matchAll(/<w:br\/>/g)].length
  const modelSoftBreaks = allInlineHolders(model).reduce(
    (n, holder) => n + holder.inlines.filter((i) => i.t === 'break').length,
    0,
  )
  if (usedBuiltinSample && modelSoftBreaks === 0) {
    problems.push('内置样本里没有软换行 —— 这一项等于没验')
  }
  if (softBreakTags !== modelSoftBreaks) {
    problems.push(
      `软换行数不符：document.xml 里 ${softBreakTags} 处 <w:br/>，模型里 ${modelSoftBreaks} 处`,
    )
  }

  /*
   * 表格：总宽/布局/行高/禁断行/整行合并/顶端对齐都必须在字节层面看得见。
   * 期望值全部从模型与规格表推导（总宽 = 版心宽、行高 = minLines × 列表段落行距），
   * 不从 Word 读回来的数推。
   */
  const tables = model.blocks.filter((b) => b.t === 'table')
  if (usedBuiltinSample && tables.length === 0) {
    problems.push('内置样本里没有表格 —— 这一项等于没验')
  }
  if (tables.length > 0) {
    // 版心宽（缇）：1in = 1440twips = 96px，即 px × 15
    const contentTwips = Math.round(
      (lengthToPx(spec.page.size.width) -
        lengthToPx(spec.page.margin.left) -
        lengthToPx(spec.page.margin.right)) *
        15,
    )
    const rowHeightPerLine = Math.round(spec.styles.listItem.linePt * 20)
    const tblXmls = [...docXml.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].map((m) => m[0])
    if (tblXmls.length !== tables.length) {
      problems.push(`表格数不符：document.xml 里 ${tblXmls.length} 张，模型里 ${tables.length} 张`)
    }
    if (/<w:tblW w:type="auto"/.test(docXml)) {
      problems.push('出现了 w:tblW w:type="auto"（表格宽度没显式写死成 dxa）')
    }

    tables.forEach((t, ti) => {
      const xml = tblXmls[ti] ?? ''
      const countIn = (re) => [...xml.matchAll(re)].length
      const rows = t.rows.length
      const mergedRows = t.rows.filter((r) => r.role !== 'body').length
      const noteRows = t.rows.filter((r) => r.role === 'note').length

      const widthTag = `<w:tblW w:type="dxa" w:w="${contentTwips}"/>`
      if (countIn(new RegExp(widthTag.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g')) !== 1) {
        problems.push(`表${ti}的总宽不是版心宽（缺 ${widthTag}）`)
      }
      if (countIn(/<w:tblLayout w:type="fixed"\/>/g) !== 1) {
        problems.push(`表${ti}不是固定布局（缺 w:tblLayout w:type="fixed"）`)
      }
      if (countIn(/<w:left w:type="dxa" w:w="108"\/>/g) !== 1 || countIn(/<w:right w:type="dxa" w:w="108"\/>/g) !== 1) {
        problems.push(`表${ti}的单元格左右内边距不是 108 缇`)
      }
      const cantSplitTag = t.cantSplit ? /<w:cantSplit\/>/g : /<w:cantSplit w:val="false"\/>/g
      if (countIn(cantSplitTag) !== rows) {
        problems.push(`表${ti}的 w:cantSplit 与模型不符（${rows} 行，模型 cantSplit=${t.cantSplit}）`)
      }
      const heightTag = new RegExp(
        `<w:trHeight w:val="${rowHeightPerLine * t.minLines}" w:hRule="atLeast"/>`,
        'g',
      )
      if (countIn(heightTag) !== rows) {
        problems.push(
          `表${ti}的行高不是 ${rowHeightPerLine * t.minLines} 缇（minLines=${t.minLines}）× ${rows} 行`,
        )
      }
      const spans = [...xml.matchAll(/<w:gridSpan w:val="(\d+)"\/>/g)].map((m) => Number(m[1]))
      if (spans.length !== mergedRows || spans.some((n) => n !== t.columns)) {
        problems.push(
          `表${ti}的整行合并不符：期望 ${mergedRows} 处 gridSpan=${t.columns}，实际 ${JSON.stringify(spans)}`,
        )
      }
      if (countIn(/<w:vAlign w:val="top"\/>/g) !== noteRows) {
        problems.push(`表${ti}的附注格顶端对齐不符（期望 ${noteRows} 处）`)
      }
    })
  }

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
  console.log(`[ok] 软换行：${softBreakTags} 处 <w:br/>（模型里 ${modelSoftBreaks} 枚）`)
  if (tables.length > 0) {
    console.log(`[ok] 表格：${tables.length} 张，总宽/固定布局/行高/禁断行/整行合并/顶端对齐均在字节层核对`)
  }
}

// 临时副本名带上 pid 与时间戳：Word 退出后会短暂占住文件，用固定名会让下一次运行
// 撞上 EBUSY —— verify:pages 会在同一个进程里连着生成两套模板的 docx，更是必须换名。
// 跑完由调用方清理。
const asciiPath = join(tmpdir(), `wtp-verify-${process.pid}-${Date.now()}.docx`)
writeFileSync(asciiPath, buffer)

const modelPath = join(root, '.qwen', 'tmp', 'model.json')
writeFileSync(modelPath, JSON.stringify({ spec, model }, null, 2), 'utf8')

console.log(
  JSON.stringify(
    {
      ok: true,
      template: template.key,
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
