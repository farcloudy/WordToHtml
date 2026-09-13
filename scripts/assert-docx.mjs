/**
 * P1 验收第 2 步：把 Word 自己报出来的属性与规格表逐项对账。
 *
 * 断言全部从 resolveSpec() 推导，不在这里重复写死数值 —— 否则改规格时
 * 测试会跟着一起错，就失去意义了。
 *
 * 用法：node scripts/assert-docx.mjs <word-dump.json> <model.json>
 */

import { readFileSync } from 'node:fs'

import {
  computeNumbering,
  lengthToPx,
  lineSpacePt,
  plainText,
  resolveSpec,
} from '../dist-lib/wordtohtml.mjs'

const [, , dumpPath, modelPath] = process.argv
if (!dumpPath || !modelPath) {
  console.error('用法: node scripts/assert-docx.mjs <word-dump.json> <model.json>')
  process.exit(2)
}

const readJson = (p) => JSON.parse(readFileSync(p, 'utf8').replace(/^\uFEFF/, ''))
const dump = readJson(dumpPath)
const { model } = readJson(modelPath)
const spec = resolveSpec()

// Word 的枚举取值
const ALIGN = { left: 0, center: 1, right: 2, both: 3 }
const LINE_RULE = { auto: 0, atLeast: 3, exact: 4 }
const WD_FIELD_PAGE = 33
const WD_REVISION_INSERT = 1
const WD_REVISION_DELETE = 2

const failures = []
const round2 = (v) => Math.round(v * 100) / 100
/** CSS 长度 → 磅。Word 的 PageSetup 全部以磅为单位。 */
const toPt = (cssLength) => (lengthToPx(cssLength) * 72) / 96

function eq(label, actual, expected) {
  if (actual === expected) return true
  failures.push(`${label} → 期望 ${JSON.stringify(expected)}，实际 ${JSON.stringify(actual)}`)
  return false
}

function near(label, actual, expected, tol = 0.05) {
  if (Math.abs(actual - expected) <= tol) return true
  failures.push(`${label} → 期望 ≈${round2(expected)}（容差 ${tol}），实际 ${actual}`)
  return false
}

function expectOneOf(label, actual, list) {
  if (list.includes(actual)) return true
  failures.push(`${label} → 期望其中之一 ${JSON.stringify(list)}，实际 ${JSON.stringify(actual)}`)
  return false
}

/* ---------------------------- 一、段落样式对账 ---------------------------- */

console.log('=== 1. 段落样式（Word 实际生效值 vs 规格表） ===')
const stylesByName = new Map(dump.styles.map((s) => [s.name, s]))

for (const kind of Object.keys(spec.styles)) {
  const s = spec.styles[kind]
  const st = stylesByName.get(s.name)
  if (!st) {
    failures.push(`样式缺失：${s.name}（对应 ${kind}）`)
    continue
  }

  const checks = [
    eq(`${s.name}·中文字体`, st.fontFarEast, s.eastAsia),
    eq(`${s.name}·西文字体`, st.fontAscii, s.ascii),
    eq(`${s.name}·字号`, st.fontSize, s.sizePt),
    eq(`${s.name}·加粗`, st.bold, s.bold ? -1 : 0),
    eq(`${s.name}·对齐`, st.alignment, ALIGN[s.align]),
    eq(`${s.name}·行距规则`, st.lineSpacingRule, LINE_RULE[s.lineRule]),
    s.lineRule === 'auto' ? true : eq(`${s.name}·行距(磅)`, st.lineSpacing, s.linePt),
    // 段前/段后的「行」以文档网格行高为基准（lineSpacePt），不是本段行距。
    // Word 报的 SpaceBefore/SpaceAfter 就是 w:before/w:after 这对后备值。
    eq(`${s.name}·段前(磅)`, st.spaceBefore, round2(lineSpacePt(s.spaceBeforeLines, spec))),
    eq(`${s.name}·段后(磅)`, st.spaceAfter, round2(lineSpacePt(s.spaceAfterLines, spec))),
    s.firstLineChars > 0
      ? eq(`${s.name}·首行缩进(字符)`, st.characterUnitFirstLineIndent, s.firstLineChars)
      : expectOneOf(`${s.name}·首行缩进(字符)`, st.characterUnitFirstLineIndent, [0, -1]),
  ]
  const ok = checks.every(Boolean)
  console.log(
    `${ok ? 'ok  ' : 'FAIL'} ${kind.padEnd(10)} ${s.name.padEnd(5)} ` +
      `${st.fontFarEast}/${st.fontAscii} ${st.fontSize}pt ` +
      `行距${st.lineSpacing}pt 段前后${st.spaceBefore}/${st.spaceAfter} ` +
      `缩进${st.characterUnitFirstLineIndent}字符 对齐${st.alignment}`,
  )
}

/* ---------------------------- 二、段落结构与文本 --------------------------- */

console.log('\n=== 2. 段落顺序、样式归属与文字（含自动编号） ===')

const numbering = computeNumbering(model.blocks, (b) =>
  b.t === 'textBlock' ? spec.styles[b.kind].numbering : 'none',
)

const expected = model.blocks
  .filter((b) => b.t === 'textBlock')
  .map((b) => ({ style: spec.styles[b.kind].name, text: (numbering.get(b.id) ?? '') + plainText(b) }))

// 换页标记在 OOXML 里都会多出一个「空段落」：
//   · 分节符由一个只带 <w:sectPr> 的空段落承载（docx 库的生成方式，副作用是每插
//     一个分节符、上一节末尾多一个空行）；
//   · 后面没有段落可挂的分页符（紧跟分节符、或在节末/文末）退回成独立段落里的
//     <w:br w:type="page"/>（见 docx/export.ts 的 sectionParagraphs）。
// Word 把这两种都读成「只含换页符（\f）的文字」，从文字上分不出来，只能按模型
// 算出应该有几个、再和 Word 报的对账。
const isBreakArtifact = (p) => p.text.length > 0 && p.text.replace(/[\f\u0007]/g, '') === ''
const artifacts = dump.paragraphs.filter(isBreakArtifact)
const body = dump.paragraphs.filter((p) => !isBreakArtifact(p))
const sectionBreaks = model.blocks.filter((b) => b.t === 'sectionBreak').length

// 一节里连续的若干分页符只产出 1 个独立段落；只有当这一串后面没有正文段落时才产出
let standalonePageBreaks = 0
let pendingPageBreak = false
for (const b of model.blocks) {
  if (b.t === 'pageBreak') {
    pendingPageBreak = true
  } else if (b.t === 'sectionBreak') {
    if (pendingPageBreak) standalonePageBreaks += 1
    pendingPageBreak = false
  } else {
    pendingPageBreak = false
  }
}
if (pendingPageBreak) standalonePageBreaks += 1

eq('换页标记承载的空段落数', artifacts.length, sectionBreaks + standalonePageBreaks)
eq('正文段落数', body.length, expected.length)
const pairs = Math.min(body.length, expected.length)
for (let i = 0; i < pairs; i += 1) {
  const got = body[i]
  const want = expected[i]
  const ok = eq(`段落${i}·样式`, got.style, want.style) && eq(`段落${i}·文字`, got.text, want.text)
  console.log(`${ok ? 'ok  ' : 'FAIL'} [${String(i).padStart(2)}] ${got.style.padEnd(5)} ${JSON.stringify(got.text)}`)
}

/* ------------------------------ 三、修订与批注 ---------------------------- */

console.log('\n=== 3. 修订与批注 ===')
eq('修订条数', dump.revisionCount, 2)
eq('批注条数', dump.commentCount, 1)
const revTypes = dump.revisions.map((r) => r.type)
expectOneOf('包含插入修订', WD_REVISION_INSERT, revTypes)
expectOneOf('包含删除修订', WD_REVISION_DELETE, revTypes)
eq('修订作者', dump.revisions[0]?.author, '张三')
const insText = dump.revisions.find((r) => r.type === WD_REVISION_INSERT)?.text ?? ''
const delText = dump.revisions.find((r) => r.type === WD_REVISION_DELETE)?.text ?? ''
eq('插入修订内容', insText, '我单位')
eq('删除修订内容', delText, '该笔债务已经清偿')
eq('批注锚定文字', dump.comments[0]?.scope ?? '', '通知书原件')
eq('批注内容', dump.comments[0]?.text ?? '', '日期需与通知书原件核对')
console.log(`ok   插入="${insText}" 删除="${delText}" 批注锚定="${dump.comments[0]?.scope}"`)

/* ------------------------------ 三b、下划线 ------------------------------- */

console.log('\n=== 3b. 下划线（w:u 是否真的生效） ===')
{
  // Word 的 Font.Underline：0 = 无，1 = 单线，9999999（wdUndefined）= 区间内混排。
  // 只对两种干净的段落表态：整段带下划线的必须报 1，完全没有下划线的必须报 0。
  // 混排的段落（同一段里既有下划线又有普通字）Word 报 wdUndefined，不在这里下结论。
  const WD_UNDERLINE_NONE = 0
  const WD_UNDERLINE_SINGLE = 1

  const blocks = model.blocks.filter((b) => b.t === 'textBlock')
  const wrong = []
  let underlined = 0
  let plain = 0
  for (let i = 0; i < pairs; i += 1) {
    const texts = (blocks[i]?.inlines ?? []).filter((x) => x.t === 'text' && x.text !== '')
    if (texts.length === 0) continue
    const got = body[i]?.underline
    if (texts.every((x) => x.underline === true)) {
      underlined += 1
      if (got !== WD_UNDERLINE_SINGLE) {
        wrong.push(`段落${i} 整段带下划线，Word 却报 Underline=${JSON.stringify(got)}`)
      }
    } else if (texts.every((x) => x.underline !== true)) {
      plain += 1
      if (got !== WD_UNDERLINE_NONE) {
        wrong.push(`段落${i} 没有下划线，Word 却报 Underline=${JSON.stringify(got)}`)
      }
    }
  }

  if (underlined === 0) wrong.push('样本里没有整段带下划线的段落 —— 这一项等于没验')
  if (wrong.length > 0) {
    for (const w of wrong) failures.push(w)
    console.log(`FAIL 下划线：${wrong.length} 处不符`)
  } else {
    console.log(
      `ok   下划线：${underlined} 段整段带下划线（Word 报 single）、` +
        `${plain} 段完全不带（Word 报 none）`,
    )
  }
}

/* ---------------------------- 四、纸张、页码、分节 ------------------------- */

console.log('\n=== 4. 纸张 / 页边距 / 页码 / 分节 ===')
eq('节数', dump.sectionCount, sectionBreaks + 1)
for (const sec of dump.sections) {
  const tag = `第${sec.index}节`
  near(`${tag}·页宽(磅)`, sec.pageWidth, toPt(spec.page.size.width), 0.1)
  near(`${tag}·页高(磅)`, sec.pageHeight, toPt(spec.page.size.height), 0.1)
  near(`${tag}·上边距(磅)`, sec.topMargin, toPt(spec.page.margin.top), 0.1)
  near(`${tag}·下边距(磅)`, sec.bottomMargin, toPt(spec.page.margin.bottom), 0.1)
  near(`${tag}·左边距(磅)`, sec.leftMargin, toPt(spec.page.margin.left), 0.1)
  near(`${tag}·右边距(磅)`, sec.rightMargin, toPt(spec.page.margin.right), 0.1)
  near(`${tag}·页脚距(磅)`, sec.footerDistance, toPt(spec.page.footer), 0.1)

  // 分节编页码用「行为」来验，而不是读属性：
  // 每一节的第一页，页脚都渲染成 "1"。如果页码没有按节重启，第 2 节会渲染成 "2"。
  // Word 对象模型的 PageNumbers.StartingNumber / RestartNumberingAtSection 在这里
  // 都返回 0，是取值怪癖（XML 里确实写了 <w:pgNumType w:start="1"/>），所以不读它。
  eq(`${tag}·页脚页码`, sec.footerText.trim(), '1')
  eq(`${tag}·页脚独立（未链接上一节）`, sec.linkedToPrevious, 0)
  expectOneOf(`${tag}·页脚含 PAGE 域`, WD_FIELD_PAGE, sec.footerFieldTypes)
  console.log(
    `ok   ${tag} ${round2(sec.pageWidth)}×${round2(sec.pageHeight)}磅 ` +
      `边距${round2(sec.topMargin)}/${round2(sec.bottomMargin)}/${round2(sec.leftMargin)}/${round2(sec.rightMargin)} ` +
      `页脚页码"${sec.footerText.trim()}" 域${JSON.stringify(sec.footerFieldTypes)}`,
  )
}

/* --------------------------------- 结论 --------------------------------- */

console.log('')
if (failures.length > 0) {
  console.error(`[FAIL] 共 ${failures.length} 项不符：`)
  for (const f of failures) console.error(`  - ${f}`)
  process.exit(1)
}
console.log('[PASS] Word 实际生效的样式、结构、修订、批注、纸张与页码全部与规格表一致。')
