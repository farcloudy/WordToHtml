/**
 * P1 验收第 2 步：把 Word 自己报出来的属性与规格表逐项对账。
 *
 * 断言全部从 resolveSpec() 推导，不在这里重复写死数值 —— 否则改规格时
 * 测试会跟着一起错，就失去意义了。
 *
 * 用法：node scripts/assert-docx.mjs <word-dump.json> <model.json> [--template <key>]
 *
 * 规格表**按 --template 的 key 从 DOC_TEMPLATES 现推**，不从 model.json 里读：
 * model.json 是上一步自己写出来的，拿它当期望值等于自己给自己判卷。
 */

import { readFileSync } from 'node:fs'

import {
  DOC_TEMPLATES,
  computeNumbering,
  lengthToPx,
  lineSpacePt,
  plainText,
  resolveSections,
  resolveSpec,
} from '../dist-lib/wordtohtml.mjs'

const [, , dumpPath, modelPath] = process.argv
if (!dumpPath || !modelPath) {
  console.error('用法: node scripts/assert-docx.mjs <word-dump.json> <model.json> [--template <key>]')
  process.exit(2)
}

const templateIndex = process.argv.indexOf('--template')
const templateKey = templateIndex >= 0 ? process.argv[templateIndex + 1] : undefined
const template = templateKey
  ? DOC_TEMPLATES.find((t) => t.key === templateKey)
  : DOC_TEMPLATES[0]
if (!template) {
  console.error(
    `[FAIL] 未知模板：${templateKey}（可选：${DOC_TEMPLATES.map((t) => t.key).join(' / ')}）`,
  )
  process.exit(2)
}

const readJson = (p) => JSON.parse(readFileSync(p, 'utf8').replace(/^\uFEFF/, ''))
const dump = readJson(dumpPath)
// 只取模型；期望值（规格表）由下面的 resolveSpec 现推
const { model } = readJson(modelPath)
const spec = resolveSpec(template.spec)
console.log(`=== 文件模板：${template.label}（${template.key}）===`)

// Word 的枚举取值
const ALIGN = { left: 0, center: 1, right: 2, both: 3 }
const LINE_RULE = { auto: 0, atLeast: 3, exact: 4 }
const WD_FIELD_PAGE = 33
const WD_REVISION_INSERT = 1
const WD_REVISION_DELETE = 2
// WdRowHeightRule.wdRowHeightAtLeast —— 行高「最小值」规则
const ROW_HEIGHT_AT_LEAST = 1
// WdCellVerticalAlignment.wdCellAlignVerticalTop / Center / Bottom
// （OOXML 的 w:vAlign 把「垂直居中」写作 center，Word 的枚举值是 1）
const CELL_ALIGN_VERTICAL_TOP = 0
const CELL_ALIGN_VERTICAL_CENTER = 1
const CELL_ALIGN_VERTICAL_BOTTOM = 3
// WdLineStyle.wdLineStyleNone / wdLineStyleSingle
const LINE_STYLE_NONE = 0
const LINE_STYLE_SINGLE = 1

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

/** 布尔式断言（没有「期望值」可写的那些） */
function ok(label, condition, detail = '') {
  if (condition) return true
  failures.push(`${label}${detail ? ` → ${detail}` : ''}`)
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
//
// Word 的 Paragraphs 集合会把表格单元格里的段落也走一遍（inTable=1），
// 正文这一段必须把它们排除，否则段落顺序与文字会被表格内容搅乱。
const isBreakArtifact = (p) => p.text.length > 0 && p.text.replace(/[\f\u0007]/g, '') === ''
const artifacts = dump.paragraphs.filter((p) => p.inTable === 0 && isBreakArtifact(p))
const body = dump.paragraphs.filter((p) => p.inTable === 0 && !isBreakArtifact(p))
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

/* ------------------------------- 三c、表格 -------------------------------- */

console.log('\n=== 3c. 表格（行数 / 格数 / 整行合并 / 行高规则 / 边框 / 对齐 / 总宽） ===')
{
  const modelTables = model.blocks.filter((b) => b.t === 'table')
  const dumpTables = dump.tables ?? []
  const before = failures.length
  /** 真读到过几个格的样式名 —— 一个都没读到说明这一项等于没验（check-docx.ps1 那侧没读到） */
  let cellStyleChecks = 0
  eq('表格数', dumpTables.length, modelTables.length)

  /*
   * 模型里的格文字 → Word 读回来的样子。
   * 软换行（<w:br/>）在 Word 的 Range.Text 里是**垂直制表符 chr(11)**，不是换行也不是空格；
   * 批注锚点不占字符。期望值必须按这个口径算，否则「格内的软换行被吞了」看不出来。
   * 格内多段落（`{p}`）：Word 的段落标记在 Range.Text 里是 **chr(13)**，check-docx.ps1
   * 只 TrimEnd 掉收尾的段落标记 / 单元格标记，中间的照原样留着，所以这里用 `\r` 拼接。
   */
  const cellText = (cell) =>
    (cell.paragraphs?.length ? cell.paragraphs : [{ inlines: [] }])
      .map((para) =>
        para.inlines
          .map((i) => (i.t === 'text' ? i.text : i.t === 'break' ? '\u000b' : ''))
          .join(''),
      )
      .join('\r')

  const softBreakInModel = model.blocks
    .filter((b) => b.t === 'table')
    .some((t) =>
      t.rows.some((r) =>
        r.cells.some((c) => (c.paragraphs ?? []).some((p) => p.inlines.some((i) => i.t === 'break'))),
      ),
    )
  if (softBreakInModel) {
    // Word 侧的直接证据：软换行在 Range.Text 里就是 chr(11)，读回来的格文字里必须看得到
    const wordSoft = dumpTables
      .flatMap((t) => (t.rows ?? []).flatMap((r) => (r.cells ?? []).map((c) => c.text ?? '')))
      .filter((text) => text.includes('\u000b')).length
    if (wordSoft === 0) {
      failures.push('模型里有软换行，Word 读回来的格文字里却没有 chr(11)（换行没落到 docx）')
    } else {
      console.log(`ok   软换行：${wordSoft} 个格子的文字里读到了 chr(11)`)
    }
  }

  /*
   * 逐格期望值的口径与 docx/export.ts 同源：
   *   水平 = cell.align.h ?? 角色默认（unit 右 / note 左 / body 跟该格样式）
   *   垂直 = cell.align.v ?? top（middle 在 OOXML 里写 center，Word 报枚举 1）
   *   样式 = 该格 kind ?? listItem；kind === 'body' 时不挂样式 → Word 报 Normal 的本地名「正文」
   *   行高 = (body 行 ? minLines : 1) × 该行各格样式 linePt 的最大值（渲染后的 0..columns-1，
   *          缺格按 listItem；表头行 / 附注行恒一行，W7）
   */
  const renderedCells = (t, row) =>
    row.role === 'body'
      ? Array.from({ length: t.columns }, (_, c) => row.cells[c] ?? { paragraphs: [{ inlines: [] }] })
      : [row.cells[0] ?? { paragraphs: [{ inlines: [] }] }]
  const alignOf = (role, cell) =>
    ALIGN[
      cell.align?.h ??
        (role === 'unit' ? 'right' : role === 'note' ? 'left' : spec.styles[cell.kind ?? 'listItem'].align)
    ]
  const vAlignOf = (cell) =>
    ({
      top: CELL_ALIGN_VERTICAL_TOP,
      middle: CELL_ALIGN_VERTICAL_CENTER,
      bottom: CELL_ALIGN_VERTICAL_BOTTOM,
    })[cell.align?.v ?? 'top']
  const styleNameOf = (cell) =>
    (cell.kind ?? 'listItem') === 'body'
      ? spec.styles.body.name
      : spec.styles[cell.kind ?? 'listItem'].name

  // 版心宽（磅）= 页面宽 − 左右页边距
  const contentWidthPt =
    ((lengthToPx(spec.page.size.width) -
      lengthToPx(spec.page.margin.left) -
      lengthToPx(spec.page.margin.right)) *
      72) /
    96

  modelTables.forEach((t, ti) => {
    const dt = dumpTables[ti]
    if (!dt) {
      failures.push(`表${ti} 在 Word 里不存在`)
      return
    }
    eq(`表${ti}·行数`, dt.rowCount, t.rows.length)
    eq(`表${ti}·列数`, dt.columnCount, t.columns)
    if (dt.widthPoints >= 0) {
      // 总宽用「第一行各格宽度之和」：Table.PreferredWidth 一旦遇到横向合并就返回
      // wdUndefined（9999999），读不出数；合并格自己的宽度报的就是整表宽。
      near(`表${ti}·总宽(磅)`, dt.widthPoints, contentWidthPt, 0.1)
    } else {
      failures.push(`表${ti} 拿不到宽度（widthPoints=-1）`)
    }

    t.rows.forEach((row, ri) => {
      const dr = dt.rows?.[ri]
      if (!dr) {
        failures.push(`表${ti} 第${ri}行在 Word 里不存在`)
        return
      }
      const merged = row.role !== 'body'
      const expectedCells = merged ? 1 : t.columns
      const tag = `表${ti}·行${ri}(${row.role})`
      eq(`${tag}·格数`, dr.cellCount, expectedCells)
      eq(`${tag}·禁止跨页断行`, dr.cantSplit, t.cantSplit ? 1 : 0)
      eq(`${tag}·行高规则`, dr.heightRule, ROW_HEIGHT_AT_LEAST)
      const cells = renderedCells(t, row)
      // 行高 = (body 行 ? minLines : 1) × 该行各格样式 linePt 的最大值（与导出侧同一条规则）。
      // 表头行 / 附注行恒一行（W7）
      const rowMaxLinePt = Math.max(...cells.map((c) => spec.styles[c.kind ?? 'listItem'].linePt))
      near(`${tag}·行高(磅)`, dr.height, (merged ? 1 : t.minLines) * rowMaxLinePt, 0.1)

      for (let c = 0; c < expectedCells; c += 1) {
        const cell = dr.cells?.[c]
        const model = cells[c] ?? { inlines: [] }
        if (!cell) {
          failures.push(`${tag} 第${c}格在 Word 里不存在`)
          continue
        }
        eq(`${tag}·格${c}文字`, cell.text, cellText(model))
        eq(`${tag}·格${c}对齐`, cell.alignment, alignOf(row.role, model))
        eq(`${tag}·格${c}合并跨度`, cell.columnSpan, merged ? t.columns : 1)
        eq(`${tag}·格${c}垂直对齐`, cell.verticalAlignment, vAlignOf(model))
        // 格内段落样式：能读就读（check-docx.ps1 读不到时留空字符串，不误报）。
        // 它读的是格内**第一段**的样式名；样式是格子级的（每个 w:p 各挂一份同一个样式），
        // 所以格内多段落时首段即代表，不需要逐段对账。
        if (typeof cell.style === 'string' && cell.style !== '') {
          cellStyleChecks += 1
          eq(`${tag}·格${c}样式`, cell.style, styleNameOf(model))
        }
        const wantStyle = merged ? LINE_STYLE_NONE : LINE_STYLE_SINGLE
        for (const side of ['top', 'left', 'bottom', 'right']) {
          eq(`${tag}·格${c}${side}线型`, cell.lineStyles?.[side], wantStyle)
        }
      }
      console.log(
        `ok   ${tag} 格数${dr.cellCount} 禁断行${dr.cantSplit} 行高${dr.heightRule}/${round2(dr.height)}磅 ` +
          `对齐[${(dr.cells ?? []).map((x) => x.alignment).join(',')}] ` +
          `垂直[${(dr.cells ?? []).map((x) => x.verticalAlignment).join(',')}]`,
      )
    })
  })

  if (modelTables.length > 0 && cellStyleChecks === 0) {
    failures.push('Word 侧一个格内样式名都没读到 —— 格内样式（w:pStyle）对账等于没验')
  }

  if (failures.length === before) {
    console.log(
      `ok   表格对账：${modelTables.length} 张表全部与模型/规格表吻合` +
        `（总宽≈${round2(contentWidthPt)}磅；逐行行高 = (body 行 ? minLines : 1) × 该行各格样式 linePt 的最大值；` +
        `逐格样式（${cellStyleChecks} 格）/水平对齐/垂直对齐逐项对上）`,
    )
  }
}

/* ---------------------------- 四、纸张、页码、分节 ------------------------- */

console.log('\n=== 4. 纸张 / 页边距 / 页码 / 分节 ===')
eq('节数', dump.sectionCount, sectionBreaks + 1)

/*
 * 逐节对账。期望值全部从 resolveSections(model, spec) 现推 —— 与导出侧、预览侧
 * 同一个真相源。W5 之前这里对**每一节**都断言「页宽=规格表、页脚文本 1、
 * linkedToPrevious=0、含 PAGE 域」，那只在「每节都独立设页脚且都重排」时才成立；
 * 现在每一节可以有各自的页面方向、页脚、页码开关，必须按它的设置算。
 *
 * 分节编号的「重排」用**行为**验，不读 PageNumbers.RestartNumberingAtSection ——
 * 那个属性在这台 Word 上恒返回 0（取值怪癖，XML 里其实写了 <w:pgNumType w:start="1"/>），
 * 「本节第一页的页脚渲染成 1」才是我们要的语义。
 */
{
  const sections = resolveSections(model, spec)
  const WD_ORIENT_PORTRAIT = 0
  const WD_ORIENT_LANDSCAPE = 1

  for (const sec of dump.sections) {
    const tag = `第${sec.index}节`
    const live = sections[sec.index - 1]
    if (!live) {
      failures.push(`${tag} 在 resolveSections 的结果里不存在（节数对不上）`)
      continue
    }
    // 横排节的纸是宽高互换过的（resolveSections 已经换过，这里不再换）
    near(`${tag}·页宽(磅)`, sec.pageWidth, toPt(live.page.size.width), 0.1)
    near(`${tag}·页高(磅)`, sec.pageHeight, toPt(live.page.size.height), 0.1)
    eq(
      `${tag}·方向`,
      sec.orientation,
      live.settings.orientation === 'landscape' ? WD_ORIENT_LANDSCAPE : WD_ORIENT_PORTRAIT,
    )
    // 页边距不随方向换（Word 的 LeftMargin 始终是「左」），四边都从规格表来
    near(`${tag}·上边距(磅)`, sec.topMargin, toPt(spec.page.margin.top), 0.1)
    near(`${tag}·下边距(磅)`, sec.bottomMargin, toPt(spec.page.margin.bottom), 0.1)
    near(`${tag}·左边距(磅)`, sec.leftMargin, toPt(spec.page.margin.left), 0.1)
    near(`${tag}·右边距(磅)`, sec.rightMargin, toPt(spec.page.margin.right), 0.1)
    near(`${tag}·页脚距(磅)`, sec.footerDistance, toPt(spec.page.footer), 0.1)

    // 关联前节 = 页脚挂在上一节上（Word 报 LinkToPrevious = 1 = True）；
    // 独立设页脚 = 0。首节没有前节，resolveSections 已把它强制成 false → 0
    eq(`${tag}·页脚是否关联前节`, sec.linkedToPrevious, live.settings.linkPrevious ? 1 : 0)

    const text = sec.footerText.trim()
    if (!live.showPageNumber) {
      eq(`${tag}·不显示页码（空页脚）`, text, '')
      eq(`${tag}·不该有 PAGE 域`, sec.footerFieldTypes.includes(WD_FIELD_PAGE), false)
    } else {
      ok(
        `${tag}·页脚渲染出页码`,
        /^\d+$/.test(text),
        `页脚文本 ${JSON.stringify(text)} 不是纯数字`,
      )
      // 首节恒从 1 开始；声明了「从 1 重排」的节，第一页就该是 1
      if (live.settings.restartAtOne || sec.index === 1) {
        eq(`${tag}·第一页页码`, text, '1')
      }
      expectOneOf(`${tag}·页脚含 PAGE 域`, WD_FIELD_PAGE, sec.footerFieldTypes)
    }
    console.log(
      `ok   ${tag} ${round2(sec.pageWidth)}×${round2(sec.pageHeight)}磅 ` +
        `方向${sec.orientation} 边距${round2(sec.topMargin)}/${round2(sec.bottomMargin)}/${round2(sec.leftMargin)}/${round2(sec.rightMargin)} ` +
        `页脚"${text}" 关联前节${sec.linkedToPrevious} 域${JSON.stringify(sec.footerFieldTypes)} ` +
        `（${live.settings.linkPrevious ? '继承前一节' : '独立页脚'}，` +
        `${live.settings.restartAtOne ? '从 1 重排' : '接着往下数'}）`,
    )
  }
}

/* --------------------------------- 结论 --------------------------------- */

console.log('')
if (failures.length > 0) {
  console.error(`[FAIL] 共 ${failures.length} 项不符：`)
  for (const f of failures) console.error(`  - ${f}`)
  process.exit(1)
}
console.log('[PASS] Word 实际生效的样式、结构、修订、批注、纸张与页码全部与规格表一致。')
