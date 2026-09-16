/**
 * 编辑操作模型层的单元测试。
 *
 * lib/edit/model.ts 是纯函数（只吃模型、不改 DOM），所以能在 node 里直接验。
 * 这里盯的是几条容易悄悄出错的地方：
 *   · 按区间替换时，夹住区间的批注锚点必须成对留下（否则渲染会把余下的文字吞进高亮）；
 *   · 切分 / 合并的边界（开头、结尾、加粗/下划线等内联格式的边界）；
 *   · 加粗、下划线、颜色的判断与增删；
 *   · 金额格式化（千分位 + 两位小数）的取舍；
 *   · 修订模式下删除不真删，而是标成 del；接受 / 拒绝修订（ins 与 del 的四种组合）；
 *   · 特殊空格（U+2003/2002/2005）能原样写进 docx 的 document.xml。
 *   · 默认快捷键表与 lib/edit/shortcuts.json 的一致性（那份 json 是 `shortcuts` prop 的默认值）。
 *
 * 用法：node scripts/test-edit-model.mjs   （需先 npm run build:lib）
 */

import { readFileSync } from 'node:fs'

import JSZip from 'jszip'

import {
  CELL_SELECTION_CLASS,
  DEFAULT_SHORTCUTS,
  SHORTCUT_ACTIONS,
  STYLE_KEYS,
  addComment,
  allInlineHolders,
  applyFormat,
  blockLength,
  bodyInsertIndex,
  bodyRowIndexes,
  buildCss,
  buildOutline,
  canJoinWithNext,
  canMergeIntoPrevious,
  cellId,
  cellParagraphCount,
  cellParagraphId,
  cellParagraphs,
  cellRectBetween,
  cellRectIndexOf,
  cellsChangingAlign,
  cellsChangingKind,
  cellsInRect,
  cellsInRects,
  cloneDoc,
  comboLabel,
  commentScopes,
  containerLength,
  contentBoxPx,
  deleteRange,
  deleteSpan,
  emptyCell,
  findBlock,
  findCell,
  findCellAt,
  findContainer,
  findMatches,
  findTable,
  formatAmount,
  hasRevisions,
  insertBodyRow,
  insertBreakAfter,
  insertColumn,
  insertSectionBreakAfter,
  insertText,
  joinWithNext,
  matchShortcut,
  mergeCellParagraph,
  mergeIntoPrevious,
  nextAlignValue,
  normalizeBlocks,
  normalizeCellCol,
  normalizeTable,
  outlineSignature,
  parseCellId,
  parseCombo,
  parseMd,
  plainText,
  rangeColor,
  rangeIsBold,
  rangeIsUnderline,
  removeBodyRow,
  removeBreak,
  removeColumn,
  removeComment,
  removeSectionBreak,
  removeTable,
  renderInlinesHtml,
  renderTableFragment,
  replyComment,
  replaceMatches,
  replaceRange,
  resolveEditorFlags,
  resolveRevisions,
  resolveSectionSettings,
  resolveSections,
  resolveShortcuts,
  resolveSpec,
  revisionSpanAt,
  sectionCountOf,
  sectionIndexOf,
  setBlockKind,
  setCellAlign,
  setCellKind,
  setCellsAlign,
  setCellsKind,
  setContainerKind,
  setMinLines,
  setRoleRow,
  setSectionSetting,
  settingsOf,
  sliceInlines,
  sortCells,
  splitBlock,
  splitCellParagraph,
  stepCell,
  storedCellAlign,
  storedCellKind,
  toBase64,
  toMd,
  updateComment,
  validateQuery,
  verticalCell,
} from '../dist-lib/wordtohtml.mjs'

let failed = 0
let passed = 0

function ok(label, condition, detail = '') {
  if (condition) {
    passed += 1
    console.log(`  ok   ${label}`)
  } else {
    failed += 1
    console.error(`  FAIL ${label}${detail ? ` — ${detail}` : ''}`)
  }
}

function eq(label, actual, expected) {
  ok(label, actual === expected, `期望 ${JSON.stringify(expected)}，实际 ${JSON.stringify(actual)}`)
}

/** 造一个只有正文段的模型 */
function docFrom(...texts) {
  return {
    blocks: texts.map((text, i) => ({
      t: 'textBlock',
      id: `t${i}`,
      kind: 'body',
      inlines: text === '' ? [] : [{ t: 'text', text }],
    })),
    comments: [],
  }
}

const textOf = (model) => model.blocks.map((b) => (b.t === 'textBlock' ? plainText(b) : '---'))

/**
 * 格内第 p 段的文字（缺省第 0 段）。
 * 格内是多段落的（`TableCellModel.paragraphs`），所以读格子文字一律按段读 ——
 * 单段格子读 `paraText(cell)` 即旧写法里的「格内文字」。
 */
const paraText = (cell, p = 0) =>
  (cellParagraphs(cell)[p]?.inlines ?? [])
    .filter((i) => i.t === 'text')
    .map((i) => i.text)
    .join('')

/** 格内各段文字，用 `|` 连起来（多段断言的便捷形式） */
const parasText = (cell) => cellParagraphs(cell).map((_, p) => paraText(cell, p)).join('|')

/** 造一张表（多段落形状）—— 章节里那些直接手搓表格的用例共用 */
const makeTable = (rows, columns = 1, minLines = 1) => ({
  t: 'table',
  id: 'tb1',
  rows,
  columns,
  minLines,
  cantSplit: true,
})

/** 一格：一段或多段（传多个字符串就是多段；不传就是一段空段） */
const cellOf = (...texts) => ({
  paragraphs: (texts.length === 0 ? [''] : texts).map((text) => ({
    inlines: text === '' ? [] : [{ t: 'text', text }],
  })),
})

const bodyRow = (...texts) => ({ role: 'body', cells: texts.map((text) => cellOf(text)) })
const roleRow = (role, text = '') => ({ role, cells: [cellOf(text)] })

console.log('=== 1. replaceRange：增 / 删 / 改 ===')
{
  const model = docFrom('abcdef')
  const block = findBlock(model, 't0')
  replaceRange(block, 3, 3, [{ t: 'text', text: 'XY' }])
  eq('插入', plainText(block), 'abcXYdef')
  replaceRange(block, 3, 5, [])
  eq('删除', plainText(block), 'abcdef')
  replaceRange(block, 0, 6, [{ t: 'text', text: '整段换掉' }])
  eq('整段替换', plainText(block), '整段换掉')
  eq('长度对得上', blockLength(block), 4)
}

console.log('\n=== 2. replaceRange：夹住区间的批注锚点必须成对留下 ===')
{
  // 模型：commentStart · "abcd" · commentEnd · "ef"
  const model = docFrom('abcdef')
  const block = findBlock(model, 't0')
  block.inlines = [
    { t: 'commentStart', commentId: 0 },
    { t: 'text', text: 'abcd' },
    { t: 'commentEnd', commentId: 0 },
    { t: 'text', text: 'ef' },
  ]
  model.comments.push({ id: 0, author: '甲', date: '2026-01-01T00:00:00.000Z', text: '注' })

  // 在批注锚定文字中间插入
  replaceRange(block, 2, 2, [{ t: 'text', text: 'ZZ' }])
  const opensAfterInsert = block.inlines.filter((i) => i.t === 'commentStart').length
  const closesAfterInsert = block.inlines.filter((i) => i.t === 'commentEnd').length
  eq('插入后开始标记仍只有一个', opensAfterInsert, 1)
  eq('插入后结束标记仍只有一个', closesAfterInsert, 1)
  eq('插入后锚定文字跟着长出来', commentScopes(model).get(0), 'abZZcd')
  eq('插入后正文正确', plainText(block), 'abZZcdef')

  // 删除跨越批注结尾的一段：此时正文是 "abZZcdef"，去掉 [3,5)
  const block2 = findBlock(model, 't0')
  replaceRange(block2, 3, 5, [])
  eq('删除后开始/结束标记仍然成对', `${block2.inlines.filter((i) => i.t === 'commentStart').length}/${block2.inlines.filter((i) => i.t === 'commentEnd').length}`, '1/1')
  eq('删除后正文正确', plainText(block2), 'abZdef')

  // 在批注起点之前插入：锚点整体后移（坐标不乱）
  const model3 = docFrom('abcdef')
  const block3 = findBlock(model3, 't0')
  block3.inlines = [
    { t: 'text', text: 'ab' },
    { t: 'commentStart', commentId: 7 },
    { t: 'text', text: 'cd' },
    { t: 'commentEnd', commentId: 7 },
    { t: 'text', text: 'ef' },
  ]
  model3.comments.push({ id: 7, author: '乙', date: '2026-01-01T00:00:00.000Z', text: '注7' })
  replaceRange(block3, 0, 0, [{ t: 'text', text: 'HEAD' }])
  eq('起点之前插入后正文正确', plainText(block3), 'HEADabcdef')
  eq('起点之前插入后锚定文字不变', commentScopes(model3).get(7), 'cd')
}

console.log('\n=== 3. splitBlock：切分与格式边界 ===')
{
  const model = docFrom('abcdef')
  const tail = splitBlock(model, 't0', 3, 'body')
  eq('切成两块', model.blocks.length, 2)
  eq('前一块', plainText(model.blocks[0]), 'abc')
  eq('后一块', plainText(model.blocks[1]), 'def')
  eq('返回新块 id', model.blocks[1].id, tail)

  const atStart = docFrom('abc')
  splitBlock(atStart, 't0', 0, 'body')
  eq('从 0 切开：首块为空', plainText(atStart.blocks[0]), '')
  eq('从 0 切开：后块是原文', plainText(atStart.blocks[1]), 'abc')

  const atEnd = docFrom('abc')
  splitBlock(atEnd, 't0', 3, 'body')
  eq('从末尾切开：尾块为空', plainText(atEnd.blocks[1]), '')

  // 加粗跨在切点上，两边各留自己那半段的格式
  const bold = {
    blocks: [
      {
        t: 'textBlock',
        id: 't0',
        kind: 'body',
        inlines: [{ t: 'text', text: 'ab', bold: true }, { t: 'text', text: 'cd' }],
      },
    ],
    comments: [],
  }
  splitBlock(bold, 't0', 2, 'body')
  eq('切点前保留加粗', bold.blocks[0].inlines.every((i) => i.bold), true)
  eq('切点后不带加粗', bold.blocks[1].inlines.every((i) => !i.bold), true)
  eq('切分后文字不变', textOf(bold).join('|'), 'ab|cd')
}

console.log('\n=== 4. mergeIntoPrevious ===')
{
  const model = docFrom('abc', 'def')
  const point = mergeIntoPrevious(model, 't1')
  eq('只剩一块', model.blocks.length, 1)
  eq('文字接上', plainText(model.blocks[0]), 'abcdef')
  eq('返回落点', `${point.blockId}@${point.offset}`, 't0@3')
  eq('首块没有上一块可并', mergeIntoPrevious(model, 't0'), null)
}

console.log('\n=== 5. 加粗判断与改色 ===')
{
  const model = {
    blocks: [
      {
        t: 'textBlock',
        id: 't0',
        kind: 'body',
        inlines: [
          { t: 'text', text: 'ab', bold: true },
          { t: 'text', text: 'cd' },
        ],
      },
    ],
    comments: [],
  }
  const block = findBlock(model, 't0')
  eq('全在加粗区间内 → true', rangeIsBold(block, 0, 2), true)
  eq('跨到不加粗区间 → false', rangeIsBold(block, 1, 3), false)
  eq('空区间 → false', rangeIsBold(block, 1, 1), false)

  applyFormat(model, 't0', 2, 4, { bold: true, color: 'FF0000' })
  eq('改完两段都加粗', block.inlines.every((i) => i.bold), true)
  eq('改完颜色一致', rangeColor(block, 2, 4), 'FF0000')
  eq('只改了一段时不一致', rangeColor(block, 0, 4), undefined)
  applyFormat(model, 't0', 0, 4, { color: null })
  eq('恢复默认色', rangeColor(block, 0, 4), undefined)
}

console.log('\n=== 5b. 下划线：增 / 删 / 与加粗互不干扰 ===')
{
  const model = docFrom('abcdef')
  const block = findBlock(model, 't0')
  eq('初始没有下划线', rangeIsUnderline(block, 0, 6), false)

  applyFormat(model, 't0', 0, 2, { underline: true })
  eq('区间内全带下划线 → true', rangeIsUnderline(block, 0, 2), true)
  eq('跨到没下划线的部分 → false', rangeIsUnderline(block, 1, 3), false)
  eq('空区间 → false', rangeIsUnderline(block, 2, 2), false)
  eq('加下划线不改文字', plainText(block), 'abcdef')
  eq('区间被切开后只落了两个字',
    block.inlines.filter((i) => i.t === 'text' && i.underline).map((i) => i.text).join(''), 'ab')
  ok('没被选中的那截不带下划线',
    block.inlines.find((i) => i.t === 'text' && i.text === 'cdef')?.underline === undefined)

  // 在下划线区间之后再套加粗：两段格式各归各的
  applyFormat(model, 't0', 2, 4, { bold: true })
  eq('后一段加粗了', rangeIsBold(block, 2, 4), true)
  eq('后一段仍没有下划线', rangeIsUnderline(block, 2, 4), false)
  eq('前一段仍只有下划线', rangeIsBold(block, 0, 2), false)

  // 改色不动下划线（applyFormat 是逐字段合并，不是整段重置）
  applyFormat(model, 't0', 0, 6, { color: 'FF0000' })
  eq('改色后下划线还在', rangeIsUnderline(block, 0, 2), true)
  eq('改色后加粗还在', rangeIsBold(block, 2, 4), true)
  eq('全文都变成红色', rangeColor(block, 0, 6), 'FF0000')

  // false 与 null 都是清除
  applyFormat(model, 't0', 0, 2, { underline: null })
  eq('null 清掉下划线', rangeIsUnderline(block, 0, 2), false)
  applyFormat(model, 't0', 2, 4, { underline: true })
  applyFormat(model, 't0', 2, 4, { underline: false })
  eq('false 也清掉下划线', rangeIsUnderline(block, 2, 4), false)
  eq('清完之后文字没变', plainText(block), 'abcdef')

  // 跨块（同一段被分页切成两片后重排回来）时按块独立判断
  const two = docFrom('上下划线', '下方没有')
  applyFormat(two, 't0', 0, 5, { underline: true })
  eq('第一段有下划线', rangeIsUnderline(findBlock(two, 't0'), 0, 5), true)
  eq('第二段不受影响', rangeIsUnderline(findBlock(two, 't1'), 0, 4), false)
}

console.log('\n=== 5c. 下划线的 md 往返 ===')
{
  const m1 = parseMd('__整段带下划线__')
  eq('解析出下划线', JSON.stringify(m1.blocks[0].inlines),
    JSON.stringify([{ t: 'text', text: '整段带下划线', underline: true }]))
  eq('序列化回同一份源码', toMd(m1), '__整段带下划线__')

  // 与加粗嵌套，两种标记都认（序列化时下划线在加粗外层，与渲染的 <u><b> 同序）
  const m2 = parseMd('**__又粗又下划线__**')
  eq('加粗与下划线能嵌套',
    JSON.stringify(m2.blocks[0].inlines),
    JSON.stringify([{ t: 'text', text: '又粗又下划线', bold: true, underline: true }]))
  eq('嵌套也能往返', toMd(m2), '__**又粗又下划线**__')
  eq('往返结构一致',
    JSON.stringify(normalizeBlocks(parseMd(toMd(m2)))),
    JSON.stringify(normalizeBlocks(m2)))

  // 不成对的 __ 是普通文字（与 ** 的既有约定一致）
  eq('落单的双下划线按原文处理', plainText(parseMd('a__b').blocks[0]), 'a__b')
  // 下划线里的文字要能带颜色
  const m3 = parseMd('{红|__红且下划线__}')
  eq('下划线能套在颜色里',
    JSON.stringify(m3.blocks[0].inlines),
    JSON.stringify([{ t: 'text', text: '红且下划线', underline: true, color: 'FF0000' }]))

  // 正文里本来就有 __ 时要转义，否则会被解析成下划线标记
  const m4 = docFrom('字段__名__')
  eq('正文里的 __ 会被转义', toMd(m4), '字段\\_\\_名\\_\\_')
  eq('转义后能原样解析回来',
    JSON.stringify(normalizeBlocks(parseMd(toMd(m4)))),
    JSON.stringify(normalizeBlocks(m4)))
}

console.log('\n=== 5d. formatAmount：千分位 + 固定两位小数 ===')
{
  const cases = [
    ['12345.6', '12,345.60'],
    ['1234', '1,234.00'],
    ['12,345.60', '12,345.60'],
    ['0', '0.00'],
    ['0.5', '0.50'],
    ['0.005', '0.01'],
    ['999.999', '1,000.00'],
    ['1234567.891', '1,234,567.89'],
    ['-1234.5', '-1,234.50'],
    ['-0.001', '0.00'],
    ['0001234', '1,234.00'],
    ['1,234,567', '1,234,567.00'],
    [' 1234.5 ', '1,234.50'],
  ]
  for (const [input, want] of cases) {
    eq(`formatAmount(${JSON.stringify(input)})`, formatAmount(input), want)
  }

  const bad = ['', '   ', 'abc', '1 234', '1.2.3', '12,34', '1e3', '￥1234', '１２３４', '1234元', '-', '.5', '1234.']
  for (const input of bad) {
    eq(`formatAmount(${JSON.stringify(input)}) → null`, formatAmount(input), null)
  }
}

console.log('\n=== 6. 批注与回复 ===')
{
  const model = docFrom('abcdef')
  const id = addComment(model, 't0', 2, 4, '这里要核', '张三', '2026-09-13T10:00:00.000Z')
  eq('批注写进模型', model.comments.length, 1)
  eq('锚定文字', commentScopes(model).get(id), 'cd')
  eq('正文没变', plainText(model.blocks[0]), 'abcdef')
  ok('批注内容对', model.comments[0].text === '这里要核' && model.comments[0].author === '张三')

  const replyId = replyComment(model, id, '已核实', '李四', '2026-09-13T11:00:00.000Z')
  eq('回复挂在父批注下', model.comments[1].parentId, id)
  eq('回复有自己的 id', replyId !== id, true)

  removeComment(model, id)
  eq('删父批注连回复一起删', model.comments.length, 0)
  ok(
    '锚点也一起清掉（只剩文字）',
    model.blocks[0].inlines.every((i) => i.t === 'text'),
    JSON.stringify(model.blocks[0].inlines),
  )
  eq('文字没被动过', plainText(model.blocks[0]), 'abcdef')
}

console.log('\n=== 7. 修订模式的插入与删除 ===')
{
  const model = docFrom('abcdef')
  insertText(model, 't0', 3, 'XY', { kind: 'ins', id: 1, author: '张三', date: '2026-09-13T10:00:00.000Z' })
  eq('插入后文字', plainText(model.blocks[0]), 'abcXYdef')
  const ins = findBlock(model, 't0').inlines.find((i) => i.rev)
  ok('插入带上了 ins 标记', ins !== undefined && ins.rev.kind === 'ins' && ins.rev.author === '张三')

  deleteRange(model, 't0', 0, 3, { kind: 'del', id: 2, author: '张三', date: '2026-09-13T10:00:00.000Z' })
  eq('删除不真删：文字还在', plainText(model.blocks[0]), 'abcXYdef')
  const del = findBlock(model, 't0').inlines.filter((i) => i.rev && i.rev.kind === 'del')
  eq('被删的三段都标了 del', del.length, 1)
  eq('标 del 的正是前三字', del[0].text, 'abc')

  const hard = docFrom('abcdef')
  deleteRange(hard, 't0', 1, 3)
  eq('没有修订标记时是真删', plainText(findBlock(hard, 't0')), 'adef')
}

console.log('\n=== 7b. 接受 / 拒绝修订（resolveRevisions / hasRevisions / revisionSpanAt）===')
{
  const rev = (kind, id) => ({ kind, id, author: '张三', date: '2026-09-13T10:00:00.000Z' })
  /** 'ab' + 插入的 'XY' + 'cdef'，再把 'a' 标成删除 —— 普通/插入/删除三段都在同一段里 */
  const tracked = () => {
    const model = docFrom('abcdef')
    insertText(model, 't0', 2, 'XY', rev('ins', 1))
    deleteRange(model, 't0', 0, 1, rev('del', 2))
    return model
  }
  /** 本段（容器）。查询类函数吃容器，改模型的那一个吃 doc + 容器 id */
  const seg = (model) => findBlock(model, 't0')
  /** 一段里的 inline 形态，带修订的写成 `文字:kind`（看「标记去掉没有」最直观） */
  const kindsOf = (model) =>
    findBlock(model, 't0')
      .inlines.map((i) => (i.rev ? `${i.text}:${i.rev.kind}` : i.text))
      .join('|')

  eq('夹具：删除、原文、插入三段挨着', kindsOf(tracked()), 'a:del|b|XY:ins|cdef')

  eq(
    '插入符落在删除修订里 → 给出整串',
    JSON.stringify(revisionSpanAt(seg(tracked()), 0)),
    '{"from":0,"to":1,"kind":"del"}',
  )
  eq(
    '插入符落在插入修订里 → 给出整串',
    JSON.stringify(revisionSpanAt(seg(tracked()), 3)),
    '{"from":2,"to":4,"kind":"ins"}',
  )
  eq('插入符落在普通文字里 → 没有', revisionSpanAt(seg(tracked()), 5), null)
  eq(
    '插入符贴在插入修订末尾（点到字右半边）也算落在里面',
    JSON.stringify(revisionSpanAt(seg(tracked()), 4)),
    '{"from":2,"to":4,"kind":"ins"}',
  )

  // 连敲几个字 = 几个 inline（每笔一枚新 id），但它们是**同一处**修订，要归成一串
  const typed = docFrom('ab')
  insertText(typed, 't0', 2, 'X', rev('ins', 1))
  insertText(typed, 't0', 3, 'Y', rev('ins', 2))
  insertText(typed, 't0', 4, 'Z', rev('ins', 3))
  eq('连敲三个字是三个 inline（各带一枚 id）', kindsOf(typed), 'ab|X:ins|Y:ins|Z:ins')
  eq(
    '但它们算同一串（点一下就能整段处理）',
    JSON.stringify(revisionSpanAt(seg(typed), 3)),
    '{"from":2,"to":5,"kind":"ins"}',
  )

  // hasRevisions：「接受/拒绝修订」按钮亮不亮就看它
  const m1 = tracked()
  eq('选区里有删除修订 → true', hasRevisions(seg(m1), 0, 1), true)
  eq('选区只覆盖普通字 → false', hasRevisions(seg(m1), 4, 5), false)
  eq('选区没盖到修订 → false', hasRevisions(seg(m1), 1, 2), false)
  eq('插入符落在修订里 → true', hasRevisions(seg(m1), 3, 3), true)
  eq('插入符不在修订里 → false', hasRevisions(seg(m1), 5, 5), false)

  // 接受：删除修订连文字删掉、插入修订去掉标记留下文字
  const acc = tracked()
  eq('接受返回 true', resolveRevisions(acc, 't0', 0, 4, 'accept'), true)
  eq('接受后：被删的字没了、插入的字留下且不再带标记', kindsOf(acc), 'b|XY|cdef')
  eq('接受后再接受一次返回 false（已经没有修订了）', resolveRevisions(acc, 't0', 0, 5, 'accept'), false)

  // 拒绝：插入修订连文字删掉、删除修订去掉标记留下文字 —— 正好回到原文
  const rej = tracked()
  eq('拒绝返回 true', resolveRevisions(rej, 't0', 0, 4, 'reject'), true)
  eq('拒绝后：被删的字回来、插入的字没了', kindsOf(rej), 'a|b|cdef')
  eq('拒绝后的文字就是原文', plainText(findBlock(rej, 't0')), 'abcdef')

  // 只有插入符（没选中）时整串处理
  eq('插入符在修订里时整串接受', resolveRevisions(typed, 't0', 3, 3, 'accept'), true)
  eq('整串接受后文字一个不少、标记全消', kindsOf(typed), 'ab|X|Y|Z')

  // 只选中修订的一半：只处理这一半，另一半仍是修订
  const half = tracked()
  eq('只选插入修订的一半也能处理', resolveRevisions(half, 't0', 2, 3, 'accept'), true)
  eq('处理过的半个不再带标记、剩下一半还在', kindsOf(half), 'a:del|b|X|Y:ins|cdef')

  // 批注锚点跨过修订：处理之后锚点必须仍成对（否则渲染会把余下文字吞进高亮里）
  const withComment = tracked()
  addComment(withComment, 't0', 0, 4, '这一段', '张三', '2026-09-13T10:00:00.000Z')
  resolveRevisions(withComment, 't0', 2, 4, 'accept')
  eq('接受修订后批注锚点仍成对', commentScopes(withComment).size, 1)
}

console.log('\n=== 8. setBlockKind / cloneDoc / 往返 ===')
{
  const model = docFrom('abc')
  setBlockKind(model, 't0', 'h1')
  eq('块型改掉', findBlock(model, 't0').kind, 'h1')

  const copy = cloneDoc(model)
  findBlock(copy, 't0').inlines[0].text = '改过了'
  eq('克隆体独立于原件', plainText(findBlock(model, 't0')), 'abc')
  eq('克隆体本身能改', plainText(findBlock(copy, 't0')), '改过了')

  const md = toMd(model)
  eq('模型能序列化成 md', md, '## abc')
}

console.log('\n=== 9. 批注改写 / 换页标记 ===')
{
  const model = docFrom('abcdef')
  const id = addComment(model, 't0', 0, 2, '原内容', '张三', '2026-09-13T10:00:00.000Z')
  eq('改写返回 true', updateComment(model, id, '改过了'), true)
  eq('内容改成新的', model.comments[0].text, '改过了')
  eq('作者没被动', model.comments[0].author, '张三')
  eq('时间没被动', model.comments[0].date, '2026-09-13T10:00:00.000Z')
  eq('改不存在的批注返回 false', updateComment(model, 999, 'x'), false)

  const pageId = insertBreakAfter(model, 't0', 'page')
  eq('分页符插在指定块之后', model.blocks[1].t, 'pageBreak')
  eq('返回的就是新块 id', model.blocks[1].id, pageId)
  const secId = insertBreakAfter(model, undefined, 'section')
  const last = model.blocks[model.blocks.length - 1]
  eq('没给落点就追加到文末', last.t, 'sectionBreak')
  eq('返回的是分节符 id', last.id, secId)
  // W5：分节符只是分节符，设置挂在 doc.sections 上；新节全默认（关联前节、不重排）
  ok('新分节符带出一节全默认的设置', sectionCountOf(model) === 2 && settingsOf(model, 1) === undefined)

  eq('分页符删得掉', removeBreak(model, pageId), true)
  eq('分页符没了', model.blocks.some((b) => b.t === 'pageBreak'), false)
  eq('文字块不归 removeBreak 管', removeBreak(model, 't0'), false)

  // 附件与换页标记都要能过 md 往返
  const round = {
    blocks: [
      { t: 'textBlock', id: 'a', kind: 'attachment', inlines: [{ t: 'text', text: '附件一' }] },
      { t: 'pageBreak', id: 'p' },
      { t: 'sectionBreak', id: 's' },
    ],
    comments: [],
  }
  eq('附件/分页符/分节符都能序列化', toMd(round), '% 附件一\n===\n---')
  eq('再解析回来还是同一份结构', JSON.stringify(normalizeBlocks(parseMd(toMd(round)))),
    JSON.stringify(normalizeBlocks(round)))
}

console.log('\n=== 10. 特殊空格能原样写进 docx ===')
{
  // 三种空格都是普通文本字符，export.ts 不为它们做任何特殊处理 —— 这一项验的就是
  // 「不做特殊处理也不出事」：它们必须原样留在 document.xml 里，不能被当成空白吃掉。
  const spaces = ['\u2003', '\u2002', '\u2005']
  const model = parseMd(`前${spaces[0]}中${spaces[1]}后${spaces[2]}末`)
  const base64 = await toBase64(model, resolveSpec(), { title: '特殊空格' })
  const zip = await JSZip.loadAsync(Buffer.from(base64, 'base64'))
  const xml = await zip.file('word/document.xml').async('string')

  for (const ch of spaces) {
    ok(`U+${ch.codePointAt(0).toString(16).toUpperCase()} 写进了 document.xml`, xml.includes(ch))
  }
  ok(
    '三个空格按顺序留在同一段文字里',
    xml.includes(`前${spaces[0]}中${spaces[1]}后${spaces[2]}末`),
  )
}

console.log('\n=== 11. findMatches：字面量 / 正则 / 非法正则 / 大小写 ===')
{
  const dotted = docFrom('xa.by', 'axb')
  eq(
    '字面量模式下 . 只当普通字符',
    JSON.stringify(findMatches(dotted, 'a.b')),
    JSON.stringify([{ blockId: 't0', from: 1, to: 4 }]),
  )
  eq(
    '正则模式下 . 是通配符',
    JSON.stringify(findMatches(dotted, 'a.b', { regex: true })),
    JSON.stringify([
      { blockId: 't0', from: 1, to: 4 },
      { blockId: 't1', from: 0, to: 3 },
    ]),
  )
  eq('空查询返回 []', JSON.stringify(findMatches(dotted, '')), '[]')
  eq('匹配不到返回 []', JSON.stringify(findMatches(dotted, 'zzz')), '[]')

  eq('合法正则 validateQuery 返回 null', validateQuery('[a-z]+', true), null)
  eq('字面量模式不做正则校验', validateQuery('[', false), null)
  ok('非法正则有错误文案', typeof validateQuery('[', true) === 'string' && validateQuery('[', true).length > 0)
  eq('非法正则 findMatches 返回 []（不抛异常）', JSON.stringify(findMatches(dotted, '[', { regex: true })), '[]')

  const multi = docFrom('债务人甲，债务人乙，债务人丙')
  eq('同一段里三处都命中', findMatches(multi, '债务人').length, 3)
  const cased = docFrom('Abc abc')
  eq(
    '默认大小写敏感',
    JSON.stringify(findMatches(cased, 'abc')),
    JSON.stringify([{ blockId: 't0', from: 4, to: 7 }]),
  )
}

console.log('\n=== 12. 跨 inline 片段匹配；跳过删除修订 ===')
{
  const anchored = {
    blocks: [
      {
        t: 'textBlock',
        id: 't0',
        kind: 'body',
        inlines: [
          { t: 'text', text: 'ab' },
          { t: 'commentStart', commentId: 0 },
          { t: 'text', text: 'cd' },
          { t: 'commentEnd', commentId: 0 },
          { t: 'text', text: 'ef', bold: true },
        ],
      },
    ],
    comments: [{ id: 0, author: '甲', date: '2026-01-01T00:00:00.000Z', text: '注' }],
  }
  eq(
    '批注锚点不占字符，匹配照样跨过去',
    JSON.stringify(findMatches(anchored, 'cde')),
    JSON.stringify([{ blockId: 't0', from: 2, to: 5 }]),
  )
  eq(
    '加粗边界不挡匹配',
    JSON.stringify(findMatches(anchored, 'def')),
    JSON.stringify([{ blockId: 't0', from: 3, to: 6 }]),
  )

  // 删除修订的文字仍留在模型里，但查找不该再命中它 —— 否则反复替换会不断叠加插入
  const removed = {
    blocks: [
      {
        t: 'textBlock',
        id: 't0',
        kind: 'body',
        inlines: [
          { t: 'text', text: 'ab' },
          {
            t: 'text',
            text: 'XY',
            rev: { kind: 'del', id: 1, author: '甲', date: '2026-01-01T00:00:00.000Z' },
          },
          { t: 'text', text: 'cd' },
        ],
      },
    ],
    comments: [],
  }
  eq('被标成 del 的文字不参与查找', JSON.stringify(findMatches(removed, 'XY')), '[]')
  eq(
    'del 之前的文字照常命中',
    JSON.stringify(findMatches(removed, 'ab')),
    JSON.stringify([{ blockId: 't0', from: 0, to: 2 }]),
  )
  eq(
    'del 之后的文字坐标仍是模型坐标（跳过的两字不进坐标）',
    JSON.stringify(findMatches(removed, 'cd')),
    JSON.stringify([{ blockId: 't0', from: 4, to: 6 }]),
  )
  eq('del 把前后断开，不跨过它匹配', JSON.stringify(findMatches(removed, 'abcd')), '[]')

  const inserted = {
    blocks: [
      {
        t: 'textBlock',
        id: 't0',
        kind: 'body',
        inlines: [
          { t: 'text', text: 'ab' },
          {
            t: 'text',
            text: 'XY',
            rev: { kind: 'ins', id: 2, author: '甲', date: '2026-01-01T00:00:00.000Z' },
          },
        ],
      },
    ],
    comments: [],
  }
  eq('插入修订的文字照常参与查找', JSON.stringify(findMatches(inserted, 'abXY')), JSON.stringify([{ blockId: 't0', from: 0, to: 4 }]))
}

console.log('\n=== 13. 范围限定（不许跨区间）===')
{
  const model = docFrom('债务人甲', '债务人乙')
  eq(
    '只在给的区间里找',
    JSON.stringify(findMatches(model, '债务人', { scope: [{ blockId: 't0', from: 0, to: 3 }] })),
    JSON.stringify([{ blockId: 't0', from: 0, to: 3 }]),
  )
  eq(
    '多个区间时各自找',
    findMatches(model, '债务人', {
      scope: [
        { blockId: 't0', from: 0, to: 3 },
        { blockId: 't1', from: 0, to: 3 },
      ],
    }).length,
    2,
  )
  eq(
    '匹配必须完整落在区间内（跨出边界就不算）',
    JSON.stringify(findMatches(model, '债务人', { scope: [{ blockId: 't0', from: 1, to: 3 }] })),
    '[]',
  )
  eq(
    '区间不能跨块拼起来',
    JSON.stringify(
      findMatches(model, '人甲债务人', {
        scope: [
          { blockId: 't0', from: 0, to: 3 },
          { blockId: 't1', from: 0, to: 3 },
        ],
      }),
    ),
    '[]',
  )
  eq('空范围列表 = 一处都不命中', JSON.stringify(findMatches(model, '债务人', { scope: [] })), '[]')
}

console.log('\n=== 14. 零宽正则不会死循环 ===')
{
  const model = docFrom('aaa', 'bbb')
  eq('^ 只产生零宽匹配 → 不算命中', JSON.stringify(findMatches(model, '^', { regex: true })), '[]')
  eq(
    'a* 只取非空那一段',
    JSON.stringify(findMatches(model, 'a*', { regex: true })),
    JSON.stringify([{ blockId: 't0', from: 0, to: 3 }]),
  )
  const started = Date.now()
  const hits = findMatches(model, 'a*', { regex: true })
  eq('零宽全部替换会终止并替换掉真正的匹配', replaceMatches(model, hits, 'X'), 1)
  eq('替换后文字正确', plainText(findBlock(model, 't0')), 'X')
  ok('零宽扫描耗时正常（不死循环）', Date.now() - started < 1000)
  eq('零宽匹配集为空时替换返回 0', replaceMatches(model, [], 'X'), 0)

  const caretOnly = docFrom('bbb')
  eq('^ 之下没有可替换的区间', replaceMatches(caretOnly, findMatches(caretOnly, '^', { regex: true }), 'X'), 0)
  eq('文档没被改', plainText(findBlock(caretOnly, 't0')), 'bbb')
}

console.log('\n=== 15. replaceMatches：从后往前、格式继承、修订留痕 ===')
{
  const model = docFrom('abcabc')
  const hits = findMatches(model, 'abc')
  eq('先找出两处', hits.length, 2)
  eq('替换处数', replaceMatches(model, hits, 'X'), 2)
  eq('从后往前替换后文字正确', plainText(findBlock(model, 't0')), 'XX')

  const grow = docFrom('abXab')
  replaceMatches(grow, findMatches(grow, 'ab'), 'ZZZ')
  eq('新文字比原文长也不会串位', plainText(findBlock(grow, 't0')), 'ZZZXZZZ')

  const shrink = docFrom('1234512345')
  replaceMatches(shrink, findMatches(shrink, '12345'), '')
  eq('替换成空串等于删除', plainText(findBlock(shrink, 't0')), '')

  const multiBlock = docFrom('旧旧', '旧旧')
  eq('跨块一起替换', replaceMatches(multiBlock, findMatches(multiBlock, '旧'), '新'), 4)
  eq('第一块', plainText(findBlock(multiBlock, 't0')), '新新')
  eq('第二块', plainText(findBlock(multiBlock, 't1')), '新新')

  // 格式继承：换掉一段加粗文字不该静默丢格式
  const formatted = {
    blocks: [
      {
        t: 'textBlock',
        id: 't0',
        kind: 'body',
        inlines: [
          { t: 'text', text: 'abc', bold: true, color: 'FF0000' },
          { t: 'text', text: 'def' },
        ],
      },
    ],
    comments: [],
  }
  const spanHits = findMatches(formatted, 'bcd')
  eq('跨格式边界命中一处', spanHits.length, 1)
  replaceMatches(formatted, spanHits, 'Q')
  const spanBlock = findBlock(formatted, 't0')
  eq('替换后文字', plainText(spanBlock), 'aQef')
  const piece = spanBlock.inlines.find((i) => i.t === 'text' && i.text === 'Q')
  ok(
    '替换结果继承区间内第一个 text 片段的格式',
    piece !== undefined && piece.bold === true && piece.color === 'FF0000',
    JSON.stringify(piece ?? null),
  )
  const trailing = spanBlock.inlines.find((i) => i.t === 'text' && i.text === 'ef')
  ok('没被替换的部分保持原格式', trailing !== undefined && trailing.bold === undefined)

  // 修订模式：旧文字标 del 留着，新文字标 ins 接在后面
  const tracked = docFrom('abc')
  let seq = 0
  const makeRev = () => ({
    kind: 'ins',
    id: seq++,
    author: '张三',
    date: '2026-09-13T10:00:00.000Z',
  })
  eq('修订模式替换一处', replaceMatches(tracked, findMatches(tracked, 'b'), 'Z', makeRev), 1)
  const trackBlock = findBlock(tracked, 't0')
  eq('旧文字没被真删，新文字接在后面', plainText(trackBlock), 'abZc')
  const dels = trackBlock.inlines.filter((i) => i.t === 'text' && i.rev?.kind === 'del')
  const inss = trackBlock.inlines.filter((i) => i.t === 'text' && i.rev?.kind === 'ins')
  eq('留下一个删除标记', dels.map((i) => i.text).join(''), 'b')
  eq('留下一个插入标记', inss.map((i) => i.text).join(''), 'Z')
  eq('删除与插入用不同的修订 id', dels[0].rev.id !== inss[0].rev.id, true)
  eq('再查同一处不会再命中（替换会收敛）', JSON.stringify(findMatches(tracked, 'b')), '[]')
}

console.log('\n=== 16. buildOutline / outlineSignature ===')
{
  const model = {
    blocks: [
      { t: 'textBlock', id: 'a', kind: 'title', inlines: [{ t: 'text', text: '文件标题' }] },
      { t: 'textBlock', id: 'b', kind: 'h1', inlines: [{ t: 'text', text: '一级' }] },
      { t: 'textBlock', id: 'c', kind: 'body', inlines: [{ t: 'text', text: '正文' }] },
      { t: 'textBlock', id: 'd', kind: 'h2', inlines: [{ t: 'text', text: '二级' }] },
      { t: 'textBlock', id: 'e', kind: 'h3', inlines: [{ t: 'text', text: '三级' }] },
    ],
    comments: [],
  }
  const numbering = new Map([
    ['b', '一、'],
    ['d', '（一）'],
    ['e', '1、'],
  ])
  const entries = buildOutline(model, numbering)
  eq('只收 h1/h2/h3（title 与正文都不进）', entries.length, 3)
  eq(
    '层级、顺序与编号前缀',
    entries.map((e) => `${e.blockId}:${e.level}:${e.prefix}${e.text}`).join('|'),
    'b:1:一、一级|d:2:（一）二级|e:3:1、三级',
  )

  const noNumber = {
    blocks: [{ t: 'textBlock', id: 'h', kind: 'h1', inlines: [{ t: 'text', text: '标题' }] }],
    comments: [],
  }
  eq('没有编号时前缀是空串', buildOutline(noNumber, new Map())[0]?.prefix, '')

  const sig = outlineSignature(entries)
  eq('同一份大纲指纹相同', outlineSignature(buildOutline(model, numbering)), sig)
  eq('空大纲指纹为空串', outlineSignature([]), '')
  ok(
    '改了标题文字指纹就变',
    outlineSignature(entries.map((e, i) => (i === 0 ? { ...e, text: '改了' } : e))) !== sig,
  )
  ok(
    '编号前缀变了指纹也变',
    outlineSignature(entries.map((e, i) => (i === 0 ? { ...e, prefix: '二、' } : e))) !== sig,
  )
  ok(
    '多了/少了一条也变',
    outlineSignature(entries.slice(0, 2)) !== sig && outlineSignature([...entries, { ...entries[0], blockId: 'z' }]) !== sig,
  )
}

console.log('\n=== 17. 表格：md 往返 / cellId / cloneDoc 深拷贝 ===')
{
  const cellText = (c) => paraText(c)

  const src = [
    ':::table minLines=2',
    '> 单位：元',
    '| **项目** | **金额** |',
    '| 甲 | 1,234.00 |',
    '| 备注\\|说明 | 含\\\\反斜杠 |',
    '< 注：以上金额不含税',
    ':::',
  ].join('\n')

  const model = parseMd(src)
  const table = model.blocks[0]
  eq('表格只产出一个块', model.blocks.length, 1)
  eq('块类型', table.t, 'table')
  eq('行数', table.rows.length, 5)
  eq('角色顺序', table.rows.map((r) => r.role).join(','), 'unit,body,body,body,note')
  eq('列数 = body 行最大格数', table.columns, 2)
  eq('minLines 从围栏读出', table.minLines, 2)
  eq('cantSplit 默认 true', table.cantSplit, true)
  eq('unit 行整行一格', table.rows[0].cells.length, 1)
  eq('note 行整行一格', table.rows[4].cells.length, 1)
  eq('unit 行文字', cellText(table.rows[0].cells[0]), '单位：元')
  eq('note 行文字', cellText(table.rows[4].cells[0]), '注：以上金额不含税')
  eq('列标题加粗', table.rows[1].cells[0].paragraphs[0].inlines[0].bold, true)
  eq('格里的竖线没被当分隔符', cellText(table.rows[3].cells[0]), '备注|说明')
  eq('格里的反斜杠原样还原', cellText(table.rows[3].cells[1]), '含\\反斜杠')
  eq('单段格子的段数 = 1（不变式 paragraphs.length >= 1）', table.rows[1].cells[0].paragraphs.length, 1)

  const md2 = toMd(model)
  eq('kwarg 只写非默认值', md2.split('\n')[0], ':::table minLines=2')
  eq('序列化字节稳定（两次往返一致）', toMd(parseMd(md2)), md2)
  eq(
    '往返结构一致',
    JSON.stringify(normalizeBlocks(parseMd(md2))),
    JSON.stringify(normalizeBlocks(model)),
  )

  eq('默认 kwarg 不写出来', toMd(parseMd(':::table\n| a |\n:::')), ':::table\n| a |\n:::')
  const noSplit = parseMd(':::table cantSplit=no\n| a |\n:::')
  eq('cantSplit=no 读出来', noSplit.blocks[0].cantSplit, false)
  eq('cantSplit=no 写回去', toMd(noSplit).split('\n')[0], ':::table cantSplit=no')

  const empty = parseMd(':::table\n:::').blocks[0]
  eq('空表规范化成 1 行', empty.rows.length, 1)
  eq('空表规范化成 1 格', empty.rows[0].cells.length, 1)
  eq('空表规范化成 1 列', empty.columns, 1)
  eq('空表的格子也有一段（不变式）', empty.rows[0].cells[0].paragraphs.length, 1)

  eq('cellId 恒指第 0 段', cellId('tb1', 2, 3), 'tb1.r2c3')
  eq('cellParagraphId 第 0 段就是 cellId', cellParagraphId('tb1', 2, 3, 0), 'tb1.r2c3')
  eq('cellParagraphId 第 N 段带 .pN', cellParagraphId('tb1', 2, 3, 2), 'tb1.r2c3.p2')
  eq(
    'parseCellId 往返（含 para 缺省 0）',
    JSON.stringify(parseCellId('tb1.r2c3')),
    JSON.stringify({ tableId: 'tb1', row: 2, col: 3, para: 0 }),
  )
  eq(
    'parseCellId 认 .pN 后缀',
    JSON.stringify(parseCellId('tb1.r2c3.p5')),
    JSON.stringify({ tableId: 'tb1', row: 2, col: 3, para: 5 }),
  )
  eq(
    'parseCellId 认显式 .p0',
    JSON.stringify(parseCellId('tb1.r0c0.p0')),
    JSON.stringify({ tableId: 'tb1', row: 0, col: 0, para: 0 }),
  )
  eq('parseCellId 拒绝非单元格 id', parseCellId('b1'), null)
  eq('parseCellId 拒绝残缺的段号', parseCellId('tb1.r2c3.pX'), null)

  // cloneDoc 必须逐层新建：撤销栈与渲染快照都靠它，共享引用会被后续编辑改到
  const copy = cloneDoc(model)
  copy.blocks[0].rows[1].cells[0].paragraphs[0].inlines[0].text = '改过了'
  copy.blocks[0].rows.push({ role: 'body', cells: [emptyCell()] })
  eq('副本的行数变了', copy.blocks[0].rows.length, 6)
  eq('原件行数没变', model.blocks[0].rows.length, 5)
  eq('原件格文字没变', cellText(model.blocks[0].rows[1].cells[0]), '项目')
  eq('不共享 rows 数组', copy.blocks[0].rows === model.blocks[0].rows, false)
  eq('不共享 cells 数组', copy.blocks[0].rows[1].cells === model.blocks[0].rows[1].cells, false)
  eq(
    '不共享 inlines 数组',
    copy.blocks[0].rows[1].cells[0].paragraphs[0].inlines ===
      model.blocks[0].rows[1].cells[0].paragraphs[0].inlines,
    false,
  )
  eq(
    '不共享 paragraphs 数组',
    copy.blocks[0].rows[1].cells[0].paragraphs === model.blocks[0].rows[1].cells[0].paragraphs,
    false,
  )
}

console.log('\n=== 17a0. 重复标题行 headerRows：md kwarg / 归一化 / cloneDoc（W11）===')
{
  const one = parseMd(':::table minLines=2 headerRows=1 cantSplit=no\n| 甲 | 乙 |\n| 1 | 2 |\n:::')
  const table = one.blocks[0]
  eq('headerRows 从围栏读出', table.headerRows, 1)
  eq(
    'kwarg 顺序固定 minLines → headerRows → cantSplit（字节稳定）',
    toMd(one).split('\n')[0],
    ':::table minLines=2 headerRows=1 cantSplit=no',
  )
  eq('再往返一次字节仍稳定', toMd(parseMd(toMd(one))), toMd(one))
  eq(
    '往返结构一致（含 headerRows）',
    JSON.stringify(normalizeBlocks(parseMd(toMd(one)))),
    JSON.stringify(normalizeBlocks(one)),
  )

  // 归一化：<= 0 / 非数字 → 不落字段；超过行数 → 夹到行数
  eq('headerRows=0 不落字段', 'headerRows' in parseMd(':::table headerRows=0\n| a |\n:::').blocks[0], false)
  eq('headerRows 缺省不落字段', 'headerRows' in parseMd(':::table\n| a |\n:::').blocks[0], false)
  eq('headerRows 非数字不落字段', 'headerRows' in parseMd(':::table headerRows=x\n| a |\n:::').blocks[0], false)
  eq(
    'headerRows 超过行数时夹到行数',
    parseMd(':::table headerRows=9\n| a |\n| b |\n:::').blocks[0].headerRows,
    2,
  )
  eq(
    '夹到行数后写出来也是夹过的值',
    toMd(parseMd(':::table headerRows=9\n| a |\n| b |\n:::')).split('\n')[0],
    ':::table headerRows=2',
  )

  // 增删行之后 N 要跟着新行数走（normalizeTable 收口），不留「N > 行数」的死状态
  const grown = parseMd(':::table headerRows=2\n| a |\n| b |\n| c |\n:::')
  removeBodyRow(grown.blocks[0], 1)
  eq('删一行后 headerRows 仍合法', grown.blocks[0].headerRows, 2)
  removeBodyRow(grown.blocks[0], 1)
  eq('删到只剩一行时 headerRows 被夹成 1', grown.blocks[0].headerRows, 1)

  // cloneDoc 漏拷 headerRows 会让撤销 / 渲染快照「重复行凭空消失」
  const marked = parseMd(':::table headerRows=1\n| a |\n| b |\n:::')
  const clone = cloneDoc(marked)
  eq('cloneDoc 拷了 headerRows', clone.blocks[0].headerRows, 1)
  clone.blocks[0].headerRows = 2
  eq('副本与原件不共享这个字段', marked.blocks[0].headerRows, 1)
}

console.log('\n=== 17a. 格内多段落：md `{p}` / 段落下标 / 切分与合并 ===')
{
  // 段间用 {p}：一个格子里两段（Word 的 w:tc 里放两个 w:p）
  const two = parseMd(':::table\n| 甲{p}乙 | b |\n:::')
  const twoCell = two.blocks[0].rows[0].cells[0]
  eq('一格解析出两段', twoCell.paragraphs.length, 2)
  eq('第一段', paraText(twoCell, 0), '甲')
  eq('第二段', paraText(twoCell, 1), '乙')
  eq('隔壁格不受影响（仍是一段）', two.blocks[0].rows[0].cells[1].paragraphs.length, 1)
  eq('写回 {p}', toMd(two), ':::table\n| 甲{p}乙 | b |\n:::')
  eq('再往返一次字节稳定', toMd(parseMd(toMd(two))), toMd(two))
  eq(
    '多段往返结构一致',
    JSON.stringify(normalizeBlocks(parseMd(toMd(two)))),
    JSON.stringify(normalizeBlocks(two)),
  )

  // 边界：尾随空段 / 首段为空 —— 都必须双向可逆、字节稳定
  const tail = parseMd(':::table\n| 甲{p} |\n:::')
  eq('尾随空段：两段', tail.blocks[0].rows[0].cells[0].paragraphs.length, 2)
  eq('尾随空段：第二段是空的', paraText(tail.blocks[0].rows[0].cells[0], 1), '')
  eq('尾随空段写回去', toMd(tail), ':::table\n| 甲{p} |\n:::')
  eq('尾随空段往返字节稳定', toMd(parseMd(toMd(tail))), toMd(tail))

  const head = parseMd(':::table\n| {p}乙 |\n:::')
  eq('首段为空：两段', head.blocks[0].rows[0].cells[0].paragraphs.length, 2)
  eq('首段为空：第一段是空的', paraText(head.blocks[0].rows[0].cells[0], 0), '')
  eq('首段为空写回去', toMd(head), ':::table\n| {p}乙 |\n:::')
  eq('首段为空往返字节稳定', toMd(parseMd(toMd(head))), toMd(head))

  eq('三段的写法', parasText(parseMd(':::table\n| 甲{p}乙{p}丙 |\n:::').blocks[0].rows[0].cells[0]), '甲|乙|丙')
  eq('段内的 {br} 仍是软换行（不切段）', parseMd(':::table\n| 甲{br}乙 |\n:::').blocks[0].rows[0].cells[0].paragraphs.length, 1)
  eq(
    '段内的 {br} 写回去',
    toMd(parseMd(':::table\n| 甲{br}乙 |\n:::')),
    ':::table\n| 甲{br}乙 |\n:::',
  )

  // 与 {@…} 同现：指令的正文部分同样要按 {p} 切
  const withAttrs = parseMd(':::table\n| {@h2,center|甲{p}乙} |\n:::')
  const attrsCell = withAttrs.blocks[0].rows[0].cells[0]
  eq('{@…} 与 {p} 同现：两段', attrsCell.paragraphs.length, 2)
  eq('{@…} 与 {p} 同现：样式落在格子上', attrsCell.kind, 'h2')
  eq('{@…} 与 {p} 同现：对齐也落在格子上', attrsCell.align?.h, 'center')
  eq('{@…} 与 {p} 同现：文字', parasText(attrsCell), '甲|乙')
  eq('{@…} 与 {p} 同现写回去', toMd(withAttrs), ':::table\n| {@h2,center|甲{p}乙} |\n:::')
  eq('{@…} 与 {p} 同现往返稳定', toMd(parseMd(toMd(withAttrs))), toMd(withAttrs))

  // 转义：正文里真写 {p} 字面量要可逆，且不许被当成段落标记
  const literal = parseMd(':::table\n| \\{p\\} |\n:::')
  const literalCell = literal.blocks[0].rows[0].cells[0]
  eq('转义后的 {p} 是字面量（不切段）', literalCell.paragraphs.length, 1)
  eq('转义后的 {p} 文字', paraText(literalCell), '{p}')
  eq('字面量写回去仍是转义形态', toMd(literal), ':::table\n| \\{p\\} |\n:::')
  eq('字面量往返稳定', toMd(parseMd(toMd(literal))), toMd(literal))

  // 段落标记只在顶层认：{} 里面、[[]] 里面的 {p} 是内容
  eq(
    '竖线指令里的 {p} 不算段落标记',
    parseMd(':::table\n| {红|甲{p}乙} |\n:::').blocks[0].rows[0].cells[0].paragraphs.length,
    1,
  )

  // splitCellParagraph / mergeCellParagraph：边界与坐标
  const doc = parseMd(':::table\n| 甲乙丙 | d |\n| e | f |\n:::')
  const table = doc.blocks[0]
  const id0 = cellId(table.id, 0, 0)
  eq('切分：返回新段落下标', splitCellParagraph(doc, id0, 1, 'listItem'), 1)
  eq('切分：格内两段', cellParagraphCount(table.rows[0].cells[0]), 2)
  eq('切分：前段', paraText(table.rows[0].cells[0], 0), '甲')
  eq('切分：后段', paraText(table.rows[0].cells[0], 1), '乙丙')
  eq('切分：后续段整体后移（第 0 段仍是甲）', paraText(table.rows[0].cells[0], 0), '甲')

  // 在第 1 段里再切一刀：新段落是 .p2 形态
  eq('再切一刀返回 2', splitCellParagraph(doc, cellParagraphId(table.id, 0, 0, 1), 1, 'listItem'), 2)
  eq('三段', parasText(table.rows[0].cells[0]), '甲|乙|丙')
  eq('第 0 段之前的段落标记不存在', splitCellParagraph(doc, cellId(table.id, 9, 9), 0, 'listItem'), null)
  eq('越界段号切不动', splitCellParagraph(doc, cellParagraphId(table.id, 0, 0, 9), 0, 'listItem'), null)
  eq('段落 id 不是格子 → null', splitCellParagraph(doc, doc.blocks[0].id, 0, 'listItem'), null)

  eq('合并点 = 前段原长度', mergeCellParagraph(doc, id0, 1), 1)
  eq('合并后两段', parasText(table.rows[0].cells[0]), '甲乙|丙')
  eq('第 0 段之前不并（不许并到上一格）', mergeCellParagraph(doc, id0, 0), null)
  eq('越界段号并不动', mergeCellParagraph(doc, id0, 9), null)
  eq('段落 id → null', mergeCellParagraph(doc, doc.blocks[0].id, 1), null)
  eq('合并到底只剩一段', (mergeCellParagraph(doc, id0, 1), parasText(table.rows[0].cells[0])), '甲乙丙')
  eq('只剩一段时再并 → null', mergeCellParagraph(doc, id0, 1), null)

  /*
   * 坏输入护栏：手搓模型（`props.model` 直给）可能缺 `paragraphs` 或给空数组。
   * 结构操作必须**先把模型补齐再改** —— 不许抛 TypeError，也不许只改一份临时兜底对象
   * （那会让函数报「切好了」而模型一字未动）。
   */
  const broken = parseMd(':::table\n| 甲 | 乙 |\n:::')
  const brokenCells = broken.blocks[0].rows[0].cells
  delete brokenCells[0].paragraphs
  brokenCells[1].paragraphs = []
  eq(
    '缺 paragraphs 的格子：切段不抛，返回新段落下标 1',
    splitCellParagraph(broken, cellId(broken.blocks[0].id, 0, 0), 0, 'listItem'),
    1,
  )
  eq('缺 paragraphs 的格子：切完模型里真的两段', brokenCells[0].paragraphs.length, 2)
  eq('空数组的格子：合并先补齐再判，返回 null', mergeCellParagraph(broken, cellId(broken.blocks[0].id, 0, 1), 1), null)
  eq('空数组的格子：补齐成一段（写回模型）', brokenCells[1].paragraphs.length, 1)
  eq(
    '空数组的格子：切段同样补齐再切',
    splitCellParagraph(broken, cellId(broken.blocks[0].id, 0, 1), 1, 'listItem'),
    1,
  )
  eq('空数组的格子：切完也是两段', brokenCells[1].paragraphs.length, 2)


  // findCellAt：格子 + 那一段一起给
  const hit = findCellAt(doc, id0)
  eq('findCellAt 拿到格子', hit?.cell === table.rows[0].cells[0], true)
  eq('findCellAt 拿到段落本尊', hit?.paragraph === table.rows[0].cells[0].paragraphs[0], true)
  eq('findCellAt 的 para', hit?.para, 0)
  eq('findCellAt 越界段 → null', findCellAt(doc, `${id0}.p9`), null)
  eq('findCellAt 越界行 → null', findCellAt(doc, cellId(table.id, 9, 0)), null)
  eq('findCellAt 认 .pN', findCellAt(doc, cellParagraphId(table.id, 0, 0, 0))?.para, 0)
}

console.log('\n=== 17b. 格内多段落：容器查找 / cloneDoc / allInlineHolders ===')
{
  const doc = parseMd(':::table\n| 甲{p}乙 | d |\n:::')
  const table = doc.blocks[0]
  const id = table.id

  // findContainer 认 .pN，且给的是那一段本尊
  eq('第 0 段', findContainer(doc, cellId(id, 0, 0)) === table.rows[0].cells[0].paragraphs[0], true)
  eq(
    '第 1 段',
    findContainer(doc, cellParagraphId(id, 0, 0, 1)) === table.rows[0].cells[0].paragraphs[1],
    true,
  )
  eq('越界段号 → undefined', findContainer(doc, cellParagraphId(id, 0, 0, 9)), undefined)
  eq('越界行号 → undefined', findContainer(doc, cellId(id, 9, 0)), undefined)
  eq('越界列号 → undefined', findContainer(doc, cellId(id, 0, 9)), undefined)
  eq('表格 id 本身不是容器', findContainer(doc, id), undefined)

  // 写入落在那一段上（replaceRange 写的是 container.inlines，包装对象写不回模型）
  insertText(doc, cellParagraphId(id, 0, 0, 1), 0, 'X')
  eq('写进第 1 段', paraText(table.rows[0].cells[0], 1), 'X乙')
  eq('第 0 段没被连带', paraText(table.rows[0].cells[0], 0), '甲')

  // cloneDoc 深拷贝：两段的 inlines 都不与原稿共享
  const copy = cloneDoc(doc)
  const src1 = table.rows[0].cells[0].paragraphs[1]
  const dst1 = copy.blocks[0].rows[0].cells[0].paragraphs[1]
  eq('第 1 段的 inlines 不共享', dst1.inlines === src1.inlines, false)
  ok('第 1 段的内容逐字拷过来', paraText(copy.blocks[0].rows[0].cells[0], 1) === 'X乙')
  dst1.inlines[0].text = '改过了'
  eq('改副本不动原稿', paraText(table.rows[0].cells[0], 1), 'X乙')

  // cloneDoc 对坏输入（格子缺 paragraphs）也要产出「恒 >= 1 段」，且不许抛
  const broken = { blocks: [{ ...makeTable([{ role: 'body', cells: [{ inlines: [{ t: 'text', text: '旧' }] }] }]) }], comments: [] }
  const fixed = cloneDoc(broken)
  eq('坏输入被补齐成一段', fixed.blocks[0].rows[0].cells[0].paragraphs.length, 1)
  eq('补出来的那一段是空的', paraText(fixed.blocks[0].rows[0].cells[0]), '')
  // 旧的 inlines 字段没有兼容层（模型里已经没有它了），克隆只保证形状合法
  eq('旧的 inlines 字段不被搬进 paragraphs', 'inlines' in fixed.blocks[0].rows[0].cells[0], false)

  // allInlineHolders 要走遍格内每一段（修订 id 扫描 / 批注 scope 都靠它）
  const holders = allInlineHolders(doc).filter((h) => h !== doc.blocks[0])
  eq('allInlineHolders 数得到格内每一段', holders.length, 3)
  eq(
    '两段都在里面',
    holders.includes(table.rows[0].cells[0].paragraphs[0]) &&
      holders.includes(table.rows[0].cells[0].paragraphs[1]),
    true,
  )

  // 批注 / 修订的扫描必须看到第 1 段里的锚点
  const marked = parseMd(':::table\n| 甲{p}[[乙|核对了]] |\n:::')
  const mt = marked.blocks[0]
  eq('第 1 段里的批注被 commentScopes 采到', commentScopes(marked).get(marked.comments[0].id), '乙')
  const rev = { t: 'text', text: '丙', rev: { kind: 'ins', id: 7, author: '甲', date: '' } }
  mt.rows[0].cells[0].paragraphs[1].inlines.push(rev)
  const found = allInlineHolders(marked)
    .flatMap((h) => h.inlines)
    .filter((i) => i.t === 'text' && i.rev)
  eq('第 1 段里的修订被扫到', found.length, 1)
  eq('修订 id 读得出来', found[0].rev.id, 7)

  // normalizeTable 补齐坏格子（缺字段 / 空数组）
  const bad = makeTable([bodyRow('a')], 1)
  bad.rows[0].cells.push({ inlines: [] }, { paragraphs: [] })
  normalizeTable(bad)
  eq('缺字段的格子被补成一段', bad.rows[0].cells[1].paragraphs.length, 1)
  eq('空数组的格子被补成一段', bad.rows[0].cells[2].paragraphs.length, 1)
}

console.log('\n=== 17c. 表格：格内行内语法 / 首尾空格 / 未闭合围栏 ===')
{
  // 这些用例是独立验收方补的：初版 splitTableCells 只看反斜杠，不跳 {} 与 [[]]，
  // 于是格内 {红|甲}、[[甲|核对原件]] 里的竖线被当成列分隔符 —— 一格拆多格、颜色与批注静默丢失。
  const cellText = (c) => paraText(c)
  const asDoc = (block) => ({ blocks: [block], comments: [] })
  const sameShape = (a, b) =>
    JSON.stringify(normalizeBlocks(a)) === JSON.stringify(normalizeBlocks(b))

  const colored = parseMd(':::table\n| {红|甲} | b |\n:::')
  eq('彩色格只算一格', colored.blocks[0].rows[0].cells.length, 2)
  eq('彩色格列数', colored.blocks[0].columns, 2)
  eq('彩色格文字', cellText(colored.blocks[0].rows[0].cells[0]), '甲')
  eq('彩色格颜色读出来', colored.blocks[0].rows[0].cells[0].paragraphs[0].inlines[0].color, 'FF0000')
  eq('彩色格往返一致', sameShape(parseMd(toMd(colored)), colored), true)

  const hex = parseMd(':::table\n| {#00FF00|乙} | c |\n:::')
  eq('十六进制色号格也只算一格', hex.blocks[0].rows[0].cells.length, 2)
  eq('十六进制色号读出来', hex.blocks[0].rows[0].cells[0].paragraphs[0].inlines[0].color, '00FF00')

  const commented = parseMd(':::table\n| [[甲|核对原件]] | d |\n:::')
  eq('批注格只算一格', commented.blocks[0].rows[0].cells.length, 2)
  eq('批注条数', commented.comments.length, 1)
  eq('批注锚定文字', cellText(commented.blocks[0].rows[0].cells[0]), '甲')
  eq('批注往返一致', sameShape(parseMd(toMd(commented)), commented), true)

  // 行内格数不齐：模型按 md 原样存，列数取 body 行的最大格数
  const ragged = parseMd(':::table\n| a | b | c |\n| d |\n:::').blocks[0]
  eq('ragged 行数', ragged.rows.length, 2)
  eq('ragged 第二行只留 1 格', ragged.rows[1].cells.length, 1)
  eq('ragged 列数 = 最大格数', ragged.columns, 3)
  eq('ragged 往返一致', sameShape(parseMd(toMd(asDoc(ragged))), asDoc(ragged)), true)

  // unit / note 行的首尾空格属于内容：不能像 body 行那样 trimEnd
  const spaced = parseMd(':::table\n>  甲 \n<  乙 \n| a |\n:::').blocks[0]
  eq('unit 行首尾空格保留', cellText(spaced.rows[0].cells[0]), ' 甲 ')
  eq('note 行首尾空格保留', cellText(spaced.rows[1].cells[0]), ' 乙 ')
  eq('空格行写回去形态一致', toMd(asDoc(spaced)).split('\n')[1], '>  甲 ')
  eq('空格往返一致', sameShape(parseMd(toMd(asDoc(spaced))), asDoc(spaced)), true)

  // 未闭合围栏：退回按普通行解，绝不把后文整篇吞掉
  const unterminated = parseMd(':::table minLines=2\n| a | b |\n\n普通一段')
  eq('未闭合围栏不吞后文', unterminated.blocks.length, 3)
  eq(
    '未闭合围栏不当表格',
    unterminated.blocks.every((b) => b.t === 'textBlock'),
    true,
  )
  eq('未闭合围栏的后文还在', plainText(unterminated.blocks[2]), '普通一段')
}

console.log('\n=== 17e. md 往返：批注内容里的 \\ 与 | 必须还原（不是越滚越多）===')
{
  // 批注内容是原样存模型、原样进 docx 的，不逐字符走 inline 解析；序列化时它会被 escapeText
  // 逃逸，解析时必须还原回来 —— 否则 `\|`、`\\` 每往返一次就多一层反斜杠（不收敛）。
  const src = '相关日期以{红|通知书}为准[[通知书原件|含\\|竖线与\\\\反斜杠]]。'
  const once = parseMd(src)
  eq('批注内容原样还原', once.comments[0].text, '含|竖线与\\反斜杠')
  const md1 = toMd(once)
  eq('写回去是逃逸形态', md1.includes('[[通知书原件|含\\|竖线与\\\\反斜杠]]'), true)
  eq('批注内容往返一次仍不变', parseMd(md1).comments[0].text, '含|竖线与\\反斜杠')
  eq('再往返一次字节仍稳定', toMd(parseMd(md1)), md1)
  const anchorText = once.blocks[0].inlines
    .filter((i) => i.t === 'text')
    .map((i) => i.text)
    .join('')
  eq('整段文字（含被锚定的那段）往返不变', anchorText, '相关日期以通知书为准通知书原件。')
}

console.log('\n=== 18. 软换行 {br}：零宽、往返、转义 ===')
{
  const doc = parseMd('甲{br}乙')
  const block = findBlock(doc, doc.blocks[0].id)
  eq('解析出 3 个 inline', block.inlines.length, 3)
  eq('中间那枚是 break', block.inlines[1].t, 'break')
  eq('零宽：plainText 不含换行', plainText(block), '甲乙')
  eq('零宽：块长度不含换行', blockLength(block), 2)
  eq('序列化写回 {br}', toMd(doc), '甲{br}乙')
  eq('往返结构一致', JSON.stringify(normalizeBlocks(parseMd(toMd(doc)))), JSON.stringify(normalizeBlocks(doc)))
  eq('再往返一次字节稳定', toMd(parseMd(toMd(doc))), '甲{br}乙')

  // 正文里真想写「{br}」时会被 escapeText 逃逸，不会撞上指令
  const literal = parseMd('\\{br\\}')
  eq('转义后的 {br} 是字面量', plainText(findBlock(literal, literal.blocks[0].id)), '{br}')
  eq('字面量往返稳定', toMd(parseMd(toMd(literal))), toMd(literal))
}

console.log('\n=== 18b. 段尾软换行的渲染：占位 <br> 只在尾随那一枚之后补 ===')
{
  const inlinesOf = (md) => {
    const doc = parseMd(md)
    return findBlock(doc, doc.blocks[0].id).inlines
  }
  eq(
    '段尾软换行之后再补一枚占位 <br>（行数才与 Word 一致）',
    renderInlinesHtml(inlinesOf('甲{br}')),
    '甲<br class="wtp-br"><br>',
  )
  eq(
    '中间的软换行不补（后面还有字，本来就有行盒）',
    renderInlinesHtml(inlinesOf('甲{br}乙')),
    '甲<br class="wtp-br">乙',
  )
  eq(
    '开头那枚不补（它后面有字）',
    renderInlinesHtml(inlinesOf('{br}甲')),
    '<br class="wtp-br">甲',
  )
  eq(
    '两枚连着结尾时补在最后一枚之后（每枚各占一行）',
    renderInlinesHtml(inlinesOf('甲{br}{br}')),
    '甲<br class="wtp-br"><br class="wtp-br"><br>',
  )
  eq('空段落仍是单枚裸 <br>（占位，不是软换行）', renderInlinesHtml([]), '<br>')
  eq(
    '格子里的尾随软换行走同一条路',
    renderTableFragment(parseMd(':::table\n| 甲{br} |\n:::').blocks[0], 0, 1).includes(
      '甲<br class="wtp-br"><br>',
    ),
    true,
  )
}

console.log('\n=== 19. 软换行的切片：不读到 undefined.length，边界不重复不丢失 ===')
{
  const inlines = [{ t: 'text', text: 'abc' }, { t: 'break' }, { t: 'text', text: 'def' }]
  const shape = (list) => list.map((i) => (i.t === 'break' ? '⏎' : i.text)).join('')
  eq('分页切片：整段保住那枚换行', shape(sliceInlines(inlines, 0, 6)), 'abc⏎def')
  eq('分页切片：前段丢掉边界上的换行', shape(sliceInlines(inlines, 0, 3)), 'abc')
  eq('分页切片：续排段不把边界换行带进来（否则多出一个没记账的行盒）', shape(sliceInlines(inlines, 3, 6)), 'def')
  eq('分页切片：换行在片内就留下', shape(sliceInlines(inlines, 0, 4)), 'abc⏎d')

  // 编辑用的 sliceStrict 与分页不同：保住比丢掉好（切分段落时不把换行弄没）
  const split = parseMd('abc{br}def')
  const host = findBlock(split, split.blocks[0].id)
  const head = host.inlines.slice(0, 1)
  eq('读第一段 inline 不崩', head.length, 1)
  splitBlock(split, host.id, 3, 'body')
  eq('在换行处分段：前段', plainText(findBlock(split, split.blocks[0].id)), 'abc')
  eq('在换行处分段：后段（换行跟着后半段）', plainText(findBlock(split, split.blocks[1].id)), 'def')
}

console.log('\n=== 20. 容器泛化：格子也能读、写、改格式 ===')
{
  const model = parseMd('正文一段\n\n:::table\n| 甲乙 | b |\n| c | d |\n:::')
  const table = model.blocks.find((b) => b.t === 'table')
  const id = table.id
  const cell = findContainer(model, cellId(id, 0, 0))
  ok('格子能被找到', cell !== undefined && cell !== null)
  eq('格子文字', containerLength(cell), 2)
  eq('表格 id 本身不是容器（表格不可编辑）', findContainer(model, id), undefined)
  eq('越界的格号不是容器', findContainer(model, cellId(id, 9, 9)), undefined)
  eq('普通段落仍能找到', findContainer(model, model.blocks[0].id) !== undefined, true)

  // 写入
  replaceRange(cell, 0, 1, [{ t: 'text', text: 'X' }])
  eq('格内替换', cell.inlines.map((i) => (i.t === 'text' ? i.text : '')).join(''), 'X乙')
  insertText(model, cellId(id, 0, 0), 1, 'Y')
  eq('格内插入', cell.inlines.map((i) => (i.t === 'text' ? i.text : '')).join(''), 'XY乙')

  // 格式：加粗 / 判读
  applyFormat(model, cellId(id, 1, 1), 0, 1, { bold: true })
  const other = findContainer(model, cellId(id, 1, 1))
  eq('格内加粗', rangeIsBold(other, 0, 1), true)
  eq('别的格子没被连带', rangeIsBold(findContainer(model, cellId(id, 1, 0)), 0, 1), false)

  // 修订模式下的删除：标成 del 留在原处
  deleteRange(model, cellId(id, 0, 1), 0, 1, {
    kind: 'del',
    id: 99,
    author: '甲',
    date: '2026-01-01T00:00:00.000Z',
  })
  const marked = findContainer(model, cellId(id, 0, 1)).inlines.find((i) => i.rev)
  eq('格内删除留痕', marked?.rev?.kind, 'del')
  eq('格内删完文字还在', containerLength(findContainer(model, cellId(id, 0, 1))), 1)

  // 结构性操作对格子必须是安全的空操作
  eq('splitBlock 对格子是空操作', splitBlock(model, cellId(id, 0, 0), 0, 'body'), '')
  eq('mergeIntoPrevious 对格子不动模型', mergeIntoPrevious(model, cellId(id, 0, 0)), null)
  eq('setBlockKind 对格子不动模型', (setBlockKind(model, cellId(id, 0, 0), 'h1'), findContainer(model, cellId(id, 0, 0)).kind), undefined)
  // joinWithNext / canJoinWithNext 同理（格子不是 textBlock，段落标记那条路不认它）
  eq('joinWithNext 对格子不动模型', joinWithNext(model, cellId(id, 0, 0)), null)
  eq('canJoinWithNext 对格子说不', canJoinWithNext(model, cellId(id, 0, 0)), false)
  eq('canMergeIntoPrevious 对格子说不', canMergeIntoPrevious(model, cellId(id, 0, 0)), false)
  eq('跨段删除不认格子端点', deleteSpan(model, { blockId: cellId(id, 0, 0), offset: 0 }, { blockId: cellId(id, 0, 0), offset: 1 }), null)
}

console.log('\n=== 21. 跨格查找与替换（含格内与段落各一处）===')
{
  const model = parseMd('债务人申报如下\n\n:::table\n| 项目 | 债务人 |\n| 甲 | 1 |\n:::')
  const table = model.blocks.find((b) => b.t === 'table')
  const hits = findMatches(model, '债务人')
  eq('段落 + 格内各命中一处', hits.length, 2)
  const inCell = hits.filter((m) => parseCellId(m.blockId) !== null)
  eq('其中一处在格子里', inCell.length, 1)
  eq('格子命中的坐标', `${inCell[0].from}/${inCell[0].to}`, '0/3')

  const replaced = replaceMatches(model, hits, '义务人')
  eq('两处都替换了', replaced, 2)
  eq('段落里换成新词', plainText(findBlock(model, model.blocks[0].id)), '义务人申报如下')
  const cell = findContainer(model, cellId(table.id, 0, 1))
  eq('格里换成新词', cell.inlines.map((i) => (i.t === 'text' ? i.text : '')).join(''), '义务人')

  // 格内多段落：每一段都要能被搜到、命中 id 是 .pN 形态、替换落在那一段上
  const multi = parseMd(':::table\n| 甲{p}债务人乙 |\n:::')
  const mhits = findMatches(multi, '债务人')
  eq('第 1 段里的字也被搜到', mhits.length, 1)
  eq('命中的 id 是 .pN 形态', mhits[0].blockId, cellParagraphId(multi.blocks[0].id, 0, 0, 1))
  eq('命中的坐标按该段算', `${mhits[0].from}/${mhits[0].to}`, '0/3')
  eq('替换写回那一段', replaceMatches(multi, mhits, '义务人'), 1)
  eq('第 0 段没被连带', paraText(multi.blocks[0].rows[0].cells[0], 0), '甲')
  eq('第 1 段换掉了', paraText(multi.blocks[0].rows[0].cells[0], 1), '义务人乙')
  eq(
    '格内多段的替换能落到 md 上',
    toMd(multi),
    ':::table\n| 甲{p}义务人乙 |\n:::',
  )
}

console.log('\n=== 22. 表格片段渲染：接口约束（外层无 data-block-id、格内坐标、行区间）===')
{
  // 这些约束是「量测与预览共用一套 DOM」的地基：外层若挂了 data-block-id，
  // fragmentOf 会把整张表当成一个片段读回模型；格内每一段若不挂，那一段就编辑不了。
  const model = parseMd(':::table minLines=2\n> 单位：元\n| 甲 | 乙 |\n| 丙{br}丁 | 戊 |\n< 注：附注\n:::')
  const table = model.blocks[0]
  const html = renderTableFragment(table, 0, table.rows.length)
  eq('外层是 wtp-table + minLines 修饰类', html.includes('<table class="wtp-table wtp-table-min2">'), true)
  eq('外层不挂 data-block-id（否则整张表会被当成片段读回）', /<table[^>]*data-block-id/.test(html), false)
  eq('外层带 data-table-id 也不用（片段自带行区间）', /data-table-id/.test(html), false)
  eq('四行各一个 <tr>', (html.match(/<tr/g) ?? []).length, 4)
  eq('unit/note 行整行一格', (html.match(/colspan="2"/g) ?? []).length, 2)
  eq('unit 行右对齐', html.includes('text-align: right'), true)
  eq('note 行左对齐且顶端对齐', html.includes('text-align: left') && html.includes('vertical-align:top'), true)
  eq('每格一层包装 wtp-cell，且包装层不挂 data-block-id', (html.match(/<div class="wtp-cell">/g) ?? []).length, 2 + 1 + 2 + 1)
  eq('包装层绝不能挂 data-block-id', /class="wtp-cell"[^>]*data-block-id/.test(html), false)
  eq('格内每一段都挂了 data-block-id', (html.match(/data-block-id="[^"]*"/g) ?? []).length, 2 + 1 + 2 + 1)
  eq(
    '段落 div 是 wtp-cellpara + wtp-<kind>（6 个格子各一段）',
    (html.match(/class="wtp-cellpara wtp-listItem"/g) ?? []).length,
    6,
  )
  eq(
    '格内坐标是 cellId 形态且 from=0',
    html.includes(`data-block-id="${cellId(table.id, 1, 0)}" data-from="0" data-to="1"`),
    true,
  )
  eq('软换行渲染成带类的 <br>（与空段落占位区分）', html.includes('<br class="wtp-br">'), true)

  // 格内多段落：每段一个 .wtp-cellpara，第 N 段的 data-block-id 带 .pN
  const multi = parseMd(':::table\n| 甲{p}乙 | 丙 |\n:::').blocks[0]
  const multiHtml = renderTableFragment(multi, 0, multi.rows.length)
  eq('一格两段 → 两个 .wtp-cellpara', (multiHtml.match(/class="wtp-cellpara/g) ?? []).length, 3)
  eq(
    '第 1 段的 data-block-id 带 .p1',
    multiHtml.includes(`data-block-id="${cellId(multi.id, 0, 0)}.p1"`),
    true,
  )
  eq(
    '第 1 段的坐标按它自己的长度算',
    multiHtml.includes(`data-block-id="${cellId(multi.id, 0, 0)}.p1" data-from="0" data-to="1"`),
    true,
  )
  eq('整格仍只挂一个 data-cell-id', (multiHtml.match(/data-cell-id="/g) ?? []).length, 2)

  const partial = renderTableFragment(table, 1, 3)
  eq('按行区间只渲那两行', (partial.match(/<tr/g) ?? []).length, 2)
  eq('行区间从 1 开始时格内 id 跟着行号走', partial.includes(`data-block-id="${cellId(table.id, 1, 0)}"`), true)

  const single = parseMd(':::table\n| a |\n:::').blocks[0]
  eq('minLines=1 用另一个修饰类', renderTableFragment(single, 0, 1).includes('wtp-table-min1'), true)
}

console.log('\n=== 23. 表格结构操作：增删行/列、unit&note 开关、归一化、行高 ===')
{
  /** 直接造一张表（不走 md，才能造出「各 body 行格数参差」这类形状） */
  // makeTable / bodyRow / roleRow 用文件开头那套（格内多段落形状）
  const shape = (table) => table.rows.map((r) => `${r.role}:${r.cells.length}`).join(',')

  // ---- insertBodyRow ----
  {
    const table = makeTable([bodyRow('a', 'b'), bodyRow('c', 'd'), bodyRow('e', 'f')], 2)
    eq('插在中间返回新行下标', insertBodyRow(table, 1), 1)
    eq('插在中间后行数 +1', table.rows.length, 4)
    eq('新行是 body', table.rows[1].role, 'body')
    eq('新行格数 = 当时的 columns', table.rows[1].cells.length, 2)
    eq('新行每格都有一段（不变式）', table.rows[1].cells.every((c) => c.paragraphs.length === 1), true)
    eq('其余行没被挪动', paraText(table.rows[2].cells[0]), 'c')

    const front = makeTable([bodyRow('a', 'b')], 2)
    eq('插在最前返回 0', insertBodyRow(front, 0), 0)
    eq('插在最前落到 0 号位', front.rows[0].role === 'body' && front.rows.length, 2)

    const end = makeTable([bodyRow('a', 'b')], 2)
    eq('插在末尾（下标 = 行数）', insertBodyRow(end, 1), 1)
    eq('插入后总数', end.rows.length, 2)

    const over = makeTable([bodyRow('a', 'b')], 2)
    eq('越界下标夹到末尾', insertBodyRow(over, 99), 1)
    eq('新行格数仍是 columns', over.rows[1].cells.length, 2)
  }

  // ---- bodyRowIndexes / bodyInsertIndex ----
  {
    const table = makeTable(
      [roleRow('unit', 'u'), bodyRow('a'), bodyRow('b'), roleRow('note', 'n')],
      1,
    )
    eq('body 行下标按显示顺序', bodyRowIndexes(table).join(','), '1,2')
    const noUnit = makeTable([bodyRow('a'), bodyRow('b')], 1)
    eq('没有 unit 时 body 下标从 0 起', bodyRowIndexes(noUnit).join(','), '0,1')
    const noBody = makeTable([roleRow('unit', 'u'), roleRow('note', 'n')], 1)
    eq('没有 body 行时是空数组', bodyRowIndexes(noBody).length, 0)

    // 光标在 unit 行时「上方插入行」的直觉落点是 0 号位，必须被夹到第一个 body 行之前，
    // 否则 body 行会插到表头行上面（note 行同理）
    eq('unit 行上「上方插入」被夹到 unit 之后', bodyInsertIndex(table, 0), 1)
    eq('unit 行上「下方插入」也落在 unit 之后', bodyInsertIndex(table, 1), 1)
    eq('note 行上「下方插入」被夹到 note 之前', bodyInsertIndex(table, table.rows.length), 3)
    eq('note 行上「上方插入」也在 note 之前', bodyInsertIndex(table, 3), 3)
    eq('body 行之间照原样', bodyInsertIndex(table, 2), 2)
    eq('负数夹到第一个 body 行', bodyInsertIndex(table, -5), 1)
    eq('越界数夹到 note 之前', bodyInsertIndex(table, 99), 3)

    const target = makeTable([roleRow('unit', 'u'), bodyRow('a'), roleRow('note', 'n')], 1)
    insertBodyRow(target, bodyInsertIndex(target, 0))
    eq('夹过之后行序仍是 unit → body → note', shape(target), 'unit:1,body:1,body:1,note:1')

    // 一个 body 行都没有的表（md 里只写了 unit / note）
    const lonely = makeTable([roleRow('unit', 'u'), roleRow('note', 'n')], 1)
    eq('无 body 行时插到 note 之前', bodyInsertIndex(lonely, 0), 1)
    const noNote = makeTable([roleRow('unit', 'u')], 1)
    eq('无 body 也无 note 时插在末尾', bodyInsertIndex(noNote, 0), 1)
    const onlyNote = makeTable([roleRow('note', 'n')], 1)
    eq('只有 note 时插在最前', bodyInsertIndex(onlyNote, 0), 0)
  }

  // ---- removeBodyRow ----
  {
    const table = makeTable([roleRow('unit', 'u'), bodyRow('a', 'b'), bodyRow('c', 'd')], 2)
    const before = JSON.stringify(table)
    eq('越界行删不掉', removeBodyRow(table, 9), false)
    eq('越界时原样不动', JSON.stringify(table), before)
    eq('非 body 行（unit）删不掉', removeBodyRow(table, 0), false)
    eq('非 body 行时原样不动', JSON.stringify(table), before)
    eq('正常删除返回 true', removeBodyRow(table, 1), true)
    eq('删完只剩 unit + 1 个 body', shape(table), 'unit:1,body:2')

    const last = makeTable([bodyRow('a', 'b')], 2)
    const lastBefore = JSON.stringify(last)
    eq('只剩一个 body 行时拒绝', removeBodyRow(last, 0), false)
    eq('拒绝时原样不动', JSON.stringify(last), lastBefore)
  }

  // ---- setRoleRow ----
  {
    const table = makeTable([bodyRow('a', 'b')], 2)
    setRoleRow(table, 'unit', true)
    eq('unit 加在最前', shape(table), 'unit:1,body:2')
    eq('unit 行只留一格', table.rows[0].cells.length, 1)
    setRoleRow(table, 'unit', true)
    eq('重复 unit=true 幂等（不产生第二行）', table.rows.filter((r) => r.role === 'unit').length, 1)

    setRoleRow(table, 'note', true)
    eq('note 加在最后', shape(table), 'unit:1,body:2,note:1')
    setRoleRow(table, 'note', true)
    eq('重复 note=true 幂等', table.rows.filter((r) => r.role === 'note').length, 1)

    setRoleRow(table, 'unit', false)
    eq('删 unit 只删它自己', shape(table), 'body:2,note:1')
    setRoleRow(table, 'note', false)
    eq('删 note 只删它自己', shape(table), 'body:2')
    setRoleRow(table, 'note', false)
    eq('note=false 且本来没有 → 空操作', shape(table), 'body:2')
  }

  // ---- insertColumn / removeColumn ----
  {
    const table = makeTable([roleRow('unit', 'u'), bodyRow('a', 'b'), bodyRow('c', 'd'), roleRow('note', 'n')], 2)
    insertColumn(table, 1)
    eq('插列只动 body 行', shape(table), 'unit:1,body:3,body:3,note:1')
    eq('unit 行格数不变', table.rows[0].cells.length, 1)
    eq('新列内容为空', table.rows[1].cells[1].paragraphs[0].inlines.length, 0)
    eq('新列也是「一段空段」（不变式）', table.rows[1].cells[1].paragraphs.length, 1)
    eq('原格向后挪', paraText(table.rows[1].cells[2]), 'b')

    // 参差行：at 按每行实际格数夹取，不要求先拍平
    const ragged = makeTable([bodyRow('a', 'b'), bodyRow('c', 'd', 'e')], 3)
    insertColumn(ragged, 2)
    eq('参差行各按自己的长度夹取', shape(ragged), 'body:3,body:4')
    eq('columns 重算为最大格数', ragged.columns, 4)

    // removeColumn：只动 body；unit/note 不变
    const del = makeTable([roleRow('unit', 'u'), bodyRow('a', 'b', 'c'), bodyRow('d', 'e', 'f'), roleRow('note', 'n')], 3)
    eq('删列成功', removeColumn(del, 1), true)
    eq('删列只动 body 行', shape(del), 'unit:1,body:2,body:2,note:1')
    eq('删掉的是第 1 列', del.rows[1].cells.map((c) => paraText(c)).join(''), 'ac')
    eq('columns 重算', del.columns, 2)

    // 最后一列不可删：逐行逐格深比较，一格都不许动
    const single = makeTable([roleRow('unit', 'u'), bodyRow('a'), roleRow('note', 'n')], 1)
    const singleBefore = JSON.stringify(single)
    eq('columns<=1 时拒绝删列', removeColumn(single, 0), false)
    eq('拒绝时逐行逐格原样', JSON.stringify(single), singleBefore)

    // 参差表删列：有的行没有这一列，就不动它
    const raggedDel = makeTable([bodyRow('a', 'b'), bodyRow('c', 'd', 'e')], 3)
    eq('参差表删列成功', removeColumn(raggedDel, 2), true)
    eq('没有该列的行不动', shape(raggedDel), 'body:2,body:2')

    /*
     * 第 0 列 + 含 unit/note 行：这是「unit/note 豁免」唯一会暴露的角落 ——
     * 下标 ≥1 时 normalizeTable 会把多出来的格裁掉，看起来像是豁免生效了；
     * 只有在下标 0 上插/删，表头与附注的文字才会被新格顶掉。所以必须专门测这一例。
     */
    const zeroCol = makeTable(
      [roleRow('unit', 'u'), bodyRow('a', 'b'), roleRow('note', 'n')],
      2,
    )
    insertColumn(zeroCol, 0)
    eq('第 0 列插列不动 unit/note', shape(zeroCol), 'unit:1,body:3,note:1')
    eq('插列后 unit 行的文字还在', paraText(zeroCol.rows[0].cells[0]), 'u')
    eq('插列后 note 行的文字还在', paraText(zeroCol.rows[2].cells[0]), 'n')
    eq('新列插在 body 行的 0 号位', zeroCol.rows[1].cells[0].paragraphs[0].inlines.length, 0)
    eq('body 行原格向后挪', paraText(zeroCol.rows[1].cells[1]), 'a')

    const zeroDel = makeTable(
      [roleRow('unit', 'u'), bodyRow('a', 'b'), roleRow('note', 'n')],
      2,
    )
    const zeroDelUnit = JSON.stringify(zeroDel.rows[0])
    const zeroDelNote = JSON.stringify(zeroDel.rows[2])
    eq('第 0 列删列成功', removeColumn(zeroDel, 0), true)
    eq('第 0 列删列不动 unit/note', shape(zeroDel), 'unit:1,body:1,note:1')
    eq('删列后 unit 行逐格原样', JSON.stringify(zeroDel.rows[0]), zeroDelUnit)
    eq('删列后 note 行逐格原样', JSON.stringify(zeroDel.rows[2]), zeroDelNote)
    eq('body 行删掉的确实是第 0 列', paraText(zeroDel.rows[1].cells[0]), 'b')
  }

  // ---- normalizeTable ----
  {
    const table = makeTable(
      [roleRow('unit', 'u'), bodyRow('a'), bodyRow('b', 'c', 'd'), roleRow('note', 'n')],
      1,
    )
    normalizeTable(table)
    eq('columns = body 行最大格数', table.columns, 3)
    eq('参差 body 行不增不减（不被拍平）', shape(table), 'unit:1,body:1,body:3,note:1')

    const trim = makeTable([roleRow('unit', 'u'), bodyRow('a', 'b', 'c')], 3)
    trim.rows[0].cells = [emptyCell(), emptyCell(), emptyCell()]
    normalizeTable(trim)
    eq('unit 行被裁到只剩第 0 格', trim.rows[0].cells.length, 1)

    const empty = makeTable([bodyRow(), bodyRow('a', 'b')], 0)
    normalizeTable(empty)
    eq('columns 取 body 行的最大值（空格不影响）', empty.columns, 2)
    const zero = makeTable([bodyRow()], 0)
    normalizeTable(zero)
    eq('全是空格时 columns 兜底为 1', zero.columns, 1)

    // 坏格子（缺 paragraphs / 空数组）要被补成一段空段，坏输入不许把后续操作带崩
    const broken = makeTable([bodyRow('a')], 1)
    broken.rows[0].cells.push({ inlines: [] }, { paragraphs: [] })
    normalizeTable(broken)
    eq('缺 paragraphs 的格子补成一段空段', broken.rows[0].cells[1].paragraphs.length, 1)
    eq('空 paragraphs 的格子补成一段空段', broken.rows[0].cells[2].paragraphs.length, 1)
    eq('补出来的那一段是空的', broken.rows[0].cells[1].paragraphs[0].inlines.length, 0)
  }

  // ---- setMinLines ----
  {
    const table = makeTable([bodyRow('a', 'b')], 2, 2)
    setMinLines(table, 1)
    eq('行高切到最小一行', table.minLines, 1)
    setMinLines(table, 2)
    eq('行高切回最小两行', table.minLines, 2)
  }

  // ---- findTable ----
  {
    const table = makeTable([bodyRow('a', 'b')], 2)
    const paragraph = { t: 'textBlock', id: 't0', kind: 'body', inlines: [{ t: 'text', text: '正文' }] }
    const model = { blocks: [paragraph, table], comments: [] }
    eq('按表格 id 命中', findTable(model, 'tb1')?.id, 'tb1')
    eq('按 cellId 命中同一张表', findTable(model, cellId('tb1', 0, 1))?.id, 'tb1')
    eq('不存在的 id → undefined', findTable(model, 'tbX'), undefined)
    eq('段落 id → undefined（它不是表）', findTable(model, 't0'), undefined)
    eq('cellId 指向的表格不存在 → undefined', findTable(model, cellId('tbX', 0, 0)), undefined)
  }
}

console.log('\n=== 24. W4b-2：键盘跨格 / 删整表 / 格内换样式 / 两组对齐 / 格首指令 ===')
{
  // makeTable / bodyRow / roleRow 用文件开头那套（格内多段落形状）
  const cellText = (c) => paraText(c)
  const step = (atom) => JSON.stringify(atom)

  // ---- A1. stepCell：行优先、unit/note 整行一格、边界 ----
  {
    const table = makeTable(
      [roleRow('unit', 'u'), bodyRow('a', 'b'), bodyRow('c', 'd'), roleRow('note', 'n')],
      2,
    )
    eq('unit 行的下一格 = 第 1 个 body 行的第 0 格', step(stepCell(table, 0, 0, 'next')), step({ row: 1, col: 0, at: 'start' }))
    eq('body 行内向右跨', step(stepCell(table, 1, 0, 'next')), step({ row: 1, col: 1, at: 'start' }))
    eq('body 行末格跨到下一行首格', step(stepCell(table, 1, 1, 'next')), step({ row: 2, col: 0, at: 'start' }))
    eq('note 行整行算一格（倒数第二格的下一格）', step(stepCell(table, 2, 1, 'next')), step({ row: 3, col: 0, at: 'start' }))
    eq('最后一格没有下一格', stepCell(table, 3, 0, 'next'), null)
    eq('Shift+Tab：从第一个 body 格回到 unit（落格尾）', step(stepCell(table, 1, 0, 'prev')), step({ row: 0, col: 0, at: 'end' }))
    eq('第一格没有上一格', stepCell(table, 0, 0, 'prev'), null)
  }

  // ---- A2. stepCell：幻影格被跳过 ----
  {
    // body 行格数少于 columns：渲染时会补空，但模型里没有这些格（findContainer 找不到）
    const ragged = makeTable([roleRow('unit', 'u'), bodyRow('a'), bodyRow('b', 'c')], 2)
    eq('幻影列（模型里没有这一格）→ null', stepCell(ragged, 1, 1, 'next'), null)
    eq('从真实末格跨到下一行首格（跳过幻影）', step(stepCell(ragged, 1, 0, 'next')), step({ row: 2, col: 0, at: 'start' }))
    eq('倒退也落在真实格上', step(stepCell(ragged, 2, 0, 'prev')), step({ row: 1, col: 0, at: 'end' }))
  }

  // ---- A3. verticalCell：同列 / 边界 / unit|note 列归 0 ----
  {
    const table = makeTable(
      [roleRow('unit', 'u'), bodyRow('a', 'b'), bodyRow('c', 'd'), roleRow('note', 'n')],
      2,
    )
    eq('同列向下', step(verticalCell(table, 1, 1, 'down')), step({ row: 2, col: 1, at: 'start' }))
    eq('同列向上', step(verticalCell(table, 2, 1, 'up')), step({ row: 1, col: 1, at: 'start' }))
    eq('向上打到 unit 行 → 列取 0', step(verticalCell(table, 1, 1, 'up')), step({ row: 0, col: 0, at: 'start' }))
    eq('向下打到 note 行 → 列取 0', step(verticalCell(table, 2, 1, 'down')), step({ row: 3, col: 0, at: 'start' }))
    eq('首行再往上没有', verticalCell(table, 0, 0, 'up'), null)
    eq('末行再往下没有', verticalCell(table, 3, 0, 'down'), null)

    const narrow = makeTable([bodyRow('a'), bodyRow('b', 'c')], 2)
    eq('目标格是幻影列 → null（fail-open）', verticalCell(narrow, 1, 1, 'up'), null)
    const short = makeTable([bodyRow('a', 'b'), bodyRow('c')], 2)
    eq('向下打到缺列的 body 行 → null', verticalCell(short, 0, 1, 'down'), null)
  }

  // ---- B. removeTable / removeBreak 收紧 ----
  {
    const docOf = () => ({
      blocks: [
        { t: 'table', id: 'tbA', rows: [bodyRow('a')], columns: 1, minLines: 1, cantSplit: true },
        { t: 'textBlock', id: 't0', kind: 'body', inlines: [] },
        { t: 'table', id: 'tbB', rows: [bodyRow('b')], columns: 1, minLines: 1, cantSplit: true },
        { t: 'textBlock', id: 't1', kind: 'body', inlines: [] },
        { t: 'table', id: 'tbC', rows: [bodyRow('c')], columns: 1, minLines: 1, cantSplit: true },
      ],
      comments: [],
    })
    const first = docOf()
    eq('删首表', removeTable(first, 'tbA'), true)
    eq('删首表后剩下的块', first.blocks.map((b) => b.id).join(','), 't0,tbB,t1,tbC')
    const middle = docOf()
    eq('删中表', removeTable(middle, 'tbB'), true)
    eq('删中表后剩下的块', middle.blocks.map((b) => b.id).join(','), 'tbA,t0,t1,tbC')
    const last = docOf()
    eq('删尾表', removeTable(last, 'tbC'), true)
    eq('删尾表后剩下的块', last.blocks.map((b) => b.id).join(','), 'tbA,t0,tbB,t1')
    const byCell = docOf()
    eq('按 cellId 也能删（复用 findTable）', removeTable(byCell, cellId('tbB', 0, 0)), true)
    eq('按 cellId 删后剩下的块', byCell.blocks.map((b) => b.id).join(','), 'tbA,t0,t1,tbC')

    const noop = docOf()
    const before = JSON.stringify(noop)
    eq('不存在的 id 是空操作', removeTable(noop, 'tbX'), false)
    eq('空操作时 doc 一个字节不变', JSON.stringify(noop), before)

    // removeBreak 收紧：拿表格 id 调它不许再删表（旧行为是个陷阱）
    const guarded = docOf()
    const guardedBefore = JSON.stringify(guarded)
    eq('removeBreak 拿表格 id 返回 false', removeBreak(guarded, 'tbB'), false)
    eq('removeBreak 拿表格 id 时表格还在且 doc 不变', JSON.stringify(guarded), guardedBefore)
    eq('removeBreak 拿段落 id 返回 false', removeBreak(guarded, 't0'), false)

    const breaks = {
      blocks: [
        { t: 'pageBreak', id: 'pg1' },
        { t: 'sectionBreak', id: 's1' },
        { t: 'textBlock', id: 't0', kind: 'body', inlines: [] },
      ],
      comments: [],
    }
    eq('分页符仍能删', removeBreak(breaks, 'pg1'), true)
    eq('分节符仍能删', removeBreak(breaks, 's1'), true)
    eq('删完只剩段落', breaks.blocks.map((b) => b.id).join(','), 't0')
  }

  // ---- C. setCellKind / setContainerKind / 渲染钩子 / md 往返 ----
  {
    const c0 = cellOf('甲')
    setCellKind(c0, 'h2')
    eq('setCellKind 设值', c0.kind, 'h2')
    setCellKind(c0, 'listItem')
    eq('回到 listItem 时删掉字段（模型不存冗余值）', 'kind' in c0, false)

    const model = parseMd('正文一段\n\n:::table\n| 甲 | 乙 |\n:::')
    const table = model.blocks.find((b) => b.t === 'table')
    setContainerKind(model, model.blocks[0].id, 'h1')
    eq('setContainerKind 改段落', findBlock(model, model.blocks[0].id).kind, 'h1')
    setContainerKind(model, cellId(table.id, 0, 0), 'h3')
    eq('setContainerKind 改格子', findCell(model, cellId(table.id, 0, 0)).kind, 'h3')
    // 带 .pN 的 id 也按「整格」处理（样式是格子级的）
    setContainerKind(model, cellParagraphId(table.id, 0, 1, 0), 'h2')
    eq('带 .p0 的 id 改的仍是整格', findCell(model, cellId(table.id, 0, 1)).kind, 'h2')

    const html = renderTableFragment(table, 0, table.rows.length)
    eq('每格一层稳定钩子 wtp-cell', (html.match(/<div class="wtp-cell">/g) ?? []).length, 2)
    eq('该格样式挂在段落 div 上（wtp-cellpara wtp-h3）', html.includes('class="wtp-cellpara wtp-h3"'), true)
    eq('钩子层与样式层是两层（包装层不带 wtp-<kind>）', /class="wtp-cell wtp-/.test(html), false)
    eq(
      '没改过的格子仍是 wtp-cellpara wtp-listItem',
      renderTableFragment(parseMd(':::table\n| 甲 | 乙 |\n:::').blocks[0], 0, 1).includes(
        'class="wtp-cellpara wtp-listItem"',
      ),
      true,
    )

    const md = toMd(model)
    eq('格内样式写进格首指令', md.includes('{@h3|甲}'), true)
    eq('格内样式往返字节稳定', toMd(parseMd(md)), md)
    eq(
      '格内样式往返结构一致',
      JSON.stringify(normalizeBlocks(parseMd(md))),
      JSON.stringify(normalizeBlocks(model)),
    )
    eq(
      'listItem 不写格首指令（老样本不变）',
      toMd({ blocks: [makeTable([bodyRow('甲')])], comments: [] }),
      ':::table\n| 甲 |\n:::',
    )
  }

  // ---- D. setCellAlign ----
  {
    const c0 = cellOf()
    setCellAlign(c0, 'h', 'center')
    eq('单维：水平', JSON.stringify(c0.align), JSON.stringify({ h: 'center' }))
    setCellAlign(c0, 'v', 'middle')
    eq('双维', JSON.stringify(c0.align), JSON.stringify({ h: 'center', v: 'middle' }))
    setCellAlign(c0, 'h', null)
    eq('清一维后另一维还在', JSON.stringify(c0.align), JSON.stringify({ v: 'middle' }))
    setCellAlign(c0, 'v', null)
    eq('两维都清掉后 align 字段消失', 'align' in c0, false)
    setCellAlign(c0, 'h', null)
    eq('本来就没有 align 时再清是空操作', 'align' in c0, false)

    const model = parseMd(':::table\n| 甲 |\n:::')
    const table = model.blocks[0]
    const mc = findCell(model, cellId(table.id, 0, 0))
    setCellAlign(mc, 'h', 'right')
    setCellAlign(mc, 'v', 'bottom')
    const md = toMd(model)
    eq('两组对齐都写进格首指令', md.includes('{@right,bottom|甲}'), true)
    eq('对齐往返字节稳定', toMd(parseMd(md)), md)
    eq(
      '对齐往返结构一致',
      JSON.stringify(normalizeBlocks(parseMd(md))),
      JSON.stringify(normalizeBlocks(model)),
    )
  }

  // ---- md：样式 + 水平 + 垂直 的综合往返；非法 token；未闭合 ----
  {
    const src = [
      ':::table minLines=2',
      '> {@h2,right|单位：元}',
      '| {@h2,center|项目} | {@center|金额} |',
      '| 甲 | {@left,bottom|1,234.00} |',
      '< {@listTitle,center|注：以上金额不含税}',
      ':::',
    ].join('\n')
    const model = parseMd(src)
    const md1 = toMd(model)
    const md2 = toMd(parseMd(md1))
    eq('综合往返：两段 md 逐字节相同', md1, md2)
    eq(
      '综合往返：结构一致',
      JSON.stringify(normalizeBlocks(parseMd(md1))),
      JSON.stringify(normalizeBlocks(model)),
    )
    eq('token 顺序固定 kind→h→v', md1.includes('{@h2,right|单位：元}'), true)
    eq('只写非默认值（纯水平）', md1.includes('{@center|金额}'), true)
    eq('两组都写', md1.includes('{@left,bottom|1,234.00}'), true)
    const t = model.blocks[0]
    eq('unit 行的 kind 读出来', t.rows[0].cells[0].kind, 'h2')
    eq('note 行的 kind 读出来', t.rows[3].cells[0].kind, 'listTitle')

    const bogus = parseMd(':::table\n| {@bogus,xx|甲} |\n:::').blocks[0]
    eq('非法 token 被忽略（文字保留）', cellText(bogus.rows[0].cells[0]), '甲')
    eq('非法 token 不产生 kind', bogus.rows[0].cells[0].kind, undefined)
    eq('非法 token 不产生 align', bogus.rows[0].cells[0].align, undefined)

    // `{@` 找不到闭合的 `}`：当普通内容，不许吞掉后面的内容
    const noClose = parseMd(':::table\n> {@h2|甲 尾\n< 注\n:::').blocks[0]
    eq('未闭合 {@：文字原样保留（不吞）', cellText(noClose.rows[0].cells[0]), '{@h2|甲 尾')
    eq('未闭合 {@：后面的行还是独立的一行', noClose.rows.length, 2)
    const noBar = parseMd(':::table\n> {@h2}甲\n:::').blocks[0]
    eq('有闭合但没分隔符 |：当普通内容，后缀不丢', cellText(noBar.rows[0].cells[0]).endsWith('甲'), true)

    // 显式写出默认样式 {@listItem|…}：解析归一化掉（模型不留冗余值，与 setCellKind 一致），
    // 序列化也不写默认值 —— 与围栏 minLines=1 / cantSplit 的写法是同一个约定
    const defModel = parseMd(':::table\n> {@listItem|甲}\n:::')
    const defCell = defModel.blocks[0].rows[0].cells[0]
    eq('显式默认样式：不落冗余 kind 字段', defCell.kind, undefined)
    eq('显式默认样式：文字保留', cellText(defCell), '甲')
    eq(
      '显式默认样式：模型→md→模型 结构一致',
      JSON.stringify(normalizeBlocks(parseMd(toMd(defModel)))),
      JSON.stringify(normalizeBlocks(defModel)),
    )
  }
}

console.log('\n=== 21. 节编辑：模型操作 / 不变式 / md 往返（W5）===')
{
  /** 「块 + 逐节设置」的完整形状（normalizeBlocks 只管 blocks，W5 起 sections 也算结构） */
  const shape = (doc) =>
    JSON.stringify({
      blocks: normalizeBlocks(doc),
      sections: doc.sections ?? null,
      editor: doc.editor ?? null,
    })

  // ---- A. sectionIndexOf：段落 / 格子里 / 分节符上 / 不存在的 id ----
  const doc = parseMd('甲\n\n---\n\n乙\n\n:::table\n| 丙 |\n:::\n\n---\n\n丁')
  const [a, s1, b, tb, , d] = doc.blocks
  eq('sectionCountOf = 分节符数 + 1', sectionCountOf(doc), 3)
  eq('首节（甲）', sectionIndexOf(doc, a.id), 0)
  eq('分节符归到它终结的那一节', sectionIndexOf(doc, s1.id), 0)
  eq('第二节（乙）', sectionIndexOf(doc, b.id), 1)
  eq('格子里也算所在节（先归到那张表）', sectionIndexOf(doc, cellId(tb.id, 0, 0)), 1)
  eq('第三节（丁）', sectionIndexOf(doc, d.id), 2)
  eq('不存在的 id 兜底为首节', sectionIndexOf(doc, 'nope'), 0)

  // ---- B. setSectionSetting：只落非默认值，无意义的组合被归一化 ----
  const m = parseMd('甲\n\n---\n\n乙\n\n---\n\n丙')
  // 「关联前节」默认是「是」，此时另两项被它接管 —— 直接设 restartAtOne 是空操作
  setSectionSetting(m, 1, { restartAtOne: true })
  eq('关联前节时设 restart 是空操作', settingsOf(m, 1), undefined)
  setSectionSetting(m, 1, { linkPrevious: false })
  eq('关联前节=否（非默认）落字段', settingsOf(m, 1)?.linkPrevious, false)
  setSectionSetting(m, 1, { restartAtOne: true })
  eq('独立后设 restart 才落字段', settingsOf(m, 1)?.restartAtOne, true)
  eq(
    '两项并存',
    JSON.stringify(settingsOf(m, 1)),
    JSON.stringify({ linkPrevious: false, restartAtOne: true }),
  )
  setSectionSetting(m, 1, { pageNumbers: true })
  eq('页码=默认值 true 不落字段', 'pageNumbers' in (settingsOf(m, 1) ?? {}), false)
  setSectionSetting(m, 1, { orientation: 'portrait' })
  eq('方向=默认值 portrait 不落字段', 'orientation' in (settingsOf(m, 1) ?? {}), false)
  setSectionSetting(m, 1, { linkPrevious: true })
  eq('关联前节回默认值 → 字段删掉', settingsOf(m, 1), undefined)
  eq('全文都回到默认 → sections 整个删掉', m.sections, undefined)

  setSectionSetting(m, 0, { linkPrevious: false })
  eq('首节的 linkPrevious 无意义 → 不落字段', m.sections, undefined)
  setSectionSetting(m, 0, { pageNumbers: false })
  eq('首节可以关掉页码', settingsOf(m, 0)?.pageNumbers, false)
  setSectionSetting(m, 99, { orientation: 'landscape' })
  eq('节号越界是空操作（不变式不被撑坏）', m.sections?.length, sectionCountOf(m))
  eq('越界写入没落进去', 'orientation' in (settingsOf(m, 99) ?? {}), false)

  const clean = parseMd('甲\n\n---\n\n乙')
  setSectionSetting(clean, 1, { linkPrevious: false })
  setSectionSetting(clean, 1, { linkPrevious: true })
  eq('全默认后 sections 整个删掉（不留等价空壳）', clean.sections, undefined)
  eq('md 也不写 kwargs', toMd(clean), '甲\n---\n乙')

  // ---- C. 增删分节符与 sections 同步（不变式：长度 = 分节符数 + 1）----
  const ins = parseMd('甲\n\n乙')
  setSectionSetting(ins, 0, { orientation: 'landscape' })
  const sId = insertSectionBreakAfter(ins, ins.blocks[0].id)
  eq('插入后 2 节', sectionCountOf(ins), 2)
  eq('sections 同步到 2 项', ins.sections?.length, 2)
  eq('不变式成立', ins.sections?.length, sectionCountOf(ins))
  eq('首节的设置没被挪走', ins.sections?.[0]?.orientation, 'landscape')
  eq('新节不沿用前一节的横向（全默认）', JSON.stringify(ins.sections?.[1]), '{}')
  eq('新分节符就在「甲」之后', ins.blocks[1].id, sId)
  eq('删分节符返回 true', removeSectionBreak(ins, sId), true)
  eq('删完 1 节', sectionCountOf(ins), 1)
  eq('sections 同步回 1 项', ins.sections?.length, 1)
  eq('删掉的是它开启的那一节（首节设置仍在）', ins.sections?.[0]?.orientation, 'landscape')
  eq('removeSectionBreak 不吃段落', removeSectionBreak(ins, ins.blocks[0].id), false)
  eq('拿不存在 id 也返回 false', removeSectionBreak(ins, 'nope'), false)

  // 中间插入：设置项要跟着下标往后挪
  const mid = parseMd('甲\n\n---\n\n乙\n\n---\n\n丙')
  setSectionSetting(mid, 2, { linkPrevious: false, restartAtOne: true })
  const newId = insertSectionBreakAfter(mid, mid.blocks[0].id)
  eq('插入后 4 节', sectionCountOf(mid), 4)
  eq('原来的第 3 节设置挪到第 4 节', mid.sections?.[3]?.restartAtOne, true)
  eq('新插入的是第 2 节（全默认）', JSON.stringify(mid.sections?.[1]), '{}')
  eq('第 1 节未受影响', JSON.stringify(mid.sections?.[0]), '{}')
  eq('删掉刚插的那个', removeSectionBreak(mid, newId), true)
  eq('设置挪回第 3 节', mid.sections?.[2]?.restartAtOne, true)
  eq('不变式仍成立', mid.sections?.length, sectionCountOf(mid))

  // ---- D. resolveSections：继承 / 方向换算 / 缺项兜底 ----
  const spec = resolveSpec()
  const rs0 = resolveSectionSettings(undefined, false)
  eq(
    '缺省 = numbers on / link on / restart off / portrait',
    JSON.stringify([rs0.pageNumbers, rs0.linkPrevious, rs0.restartAtOne, rs0.orientation]),
    JSON.stringify([true, true, false, 'portrait']),
  )
  eq('首节 linkPrevious 被强制 false', resolveSectionSettings({ linkPrevious: true }, true).linkPrevious, false)

  const rd = parseMd('甲\n\n---\n\n乙\n\n---\n\n丙')
  setSectionSetting(rd, 0, { pageNumbers: false })
  const rs1 = resolveSections(rd, spec)
  eq('解析出 3 节', rs1.length, 3)
  eq('首节显示结果 = 自己关掉了页码', rs1[0].showPageNumber, false)
  eq('第 1 节默认关联前节', rs1[1].settings.linkPrevious, true)
  eq('第 1 节继承「不显示页码」', rs1[1].showPageNumber, false)
  eq('第 2 节继续继承', rs1[2].showPageNumber, false)
  eq('纵向时纸宽 = 规格表', rs1[1].page.size.width, spec.page.size.width)
  eq('版心宽与 contentBoxPx 同源', rs1[1].content.width, contentBoxPx(spec).width)

  setSectionSetting(rd, 1, { linkPrevious: false, pageNumbers: true })
  const rs2 = resolveSections(rd, spec)
  eq('第 1 节独立后重新显示页码', rs2[1].showPageNumber, true)
  eq('第 2 节仍关联前节 → 跟着变 true', rs2[2].showPageNumber, true)

  setSectionSetting(rd, 1, { orientation: 'landscape' })
  const rs3 = resolveSections(rd, spec)
  eq('横向：纸宽 = 规格表的高', rs3[1].page.size.width, spec.page.size.height)
  eq('横向：纸高 = 规格表的宽', rs3[1].page.size.height, spec.page.size.width)
  // 四边 25mm 对称，所以横向版心高恰等于纵向版心宽
  eq('横向：版心高 = 纵向版心宽', rs3[1].content.height, contentBoxPx(spec).width)
  eq('其余节不受影响', rs3[0].page.size.width, spec.page.size.width)

  const hand = parseMd('甲\n\n---\n\n乙')
  eq('手搓模型（无 sections）也解析出 2 节', resolveSections(hand, spec).length, 2)
  eq('缺项按全默认（显示页码）', resolveSections(hand, spec)[1].showPageNumber, true)
  const short = {
    blocks: hand.blocks,
    comments: [],
    sections: [{ orientation: 'landscape' }],
  }
  const rsShort = resolveSections(short, spec)
  eq('sections 偏短也解析出 2 节', rsShort.length, 2)
  eq('偏短的第 1 节按全默认', rsShort[1].settings.restartAtOne, false)
  eq('已有的第 0 项仍生效', rsShort[0].settings.orientation, 'landscape')

  // ---- E. md 往返：非默认值收敛、字节稳定 ----
  const src = [
    '::section numbers=off',
    '甲',
    '',
    '--- link=off restart=on',
    '',
    '乙',
    '',
    '--- orientation=landscape',
    '',
    '丙',
  ].join('\n')
  const md = parseMd(src)
  eq('首节 ::section 读出来', JSON.stringify(settingsOf(md, 0)), JSON.stringify({ pageNumbers: false }))
  eq(
    '第 1 节 kwargs 读出来',
    JSON.stringify(settingsOf(md, 1)),
    JSON.stringify({ linkPrevious: false, restartAtOne: true }),
  )
  eq('第 2 节 kwargs 读出来', JSON.stringify(settingsOf(md, 2)), JSON.stringify({ orientation: 'landscape' }))
  // toMd 不写空行（本项目 md 的既有约定：空行只起分隔作用，不产出内容）
  const out = [
    '::section numbers=off',
    '甲',
    '--- link=off restart=on',
    '乙',
    '--- orientation=landscape',
    '丙',
  ].join('\n')
  eq('序列化写回 kwargs（token 顺序固定）', toMd(md), out)
  eq('再往返一次字节仍稳定', toMd(parseMd(out)), out)
  eq('往返结构一致（含 sections）', shape(parseMd(out)), shape(md))

  // 写出默认值会被归一化掉（与围栏 minLines=1 / 格内 {@listItem|…} 同一个约定）
  const obvious = parseMd('甲\n\n--- link=on numbers=on restart=off orientation=portrait\n\n乙')
  eq('显式写出的默认值不落字段', obvious.sections, undefined)
  eq('也不写回去', toMd(obvious), '甲\n---\n乙')
  const meaningless = parseMd('甲\n\n--- numbers=off\n\n乙')
  eq('关联前节时 numbers 无意义 → 不落字段', settingsOf(meaningless, 1), undefined)
  eq('也不写回去', toMd(meaningless), '甲\n---\n乙')

  // 首节指令只认文档第一处非空行，正文里的 ::section 是普通文字
  const notFirst = parseMd('甲\n\n::section numbers=off')
  eq('正文里的 ::section 不当指令', toMd(notFirst), '甲\n::section numbers=off')
  eq('也不产生 sections', notFirst.sections, undefined)
}

console.log('\n=== 25. 跨段删除：deleteSpan / joinWithNext（issues/20260915）===')
{
  // 同一段里删一段
  const one = docFrom('abcdef')
  eq('段内删除返回落点', JSON.stringify(deleteSpan(one, { blockId: 't0', offset: 2 }, { blockId: 't0', offset: 4 })), '{"blockId":"t0","offset":2}')
  eq('段内删除结果', textOf(one).join('|'), 'abef')

  // 跨段：段落标记被删掉，两头并成一段，中间整段消失
  const join = docFrom('abc', 'def', 'ghi')
  eq('跨段删除返回落点', JSON.stringify(deleteSpan(join, { blockId: 't0', offset: 1 }, { blockId: 't2', offset: 2 })), '{"blockId":"t0","offset":1}')
  eq('首尾接起来、中间段没了', textOf(join).join('|'), 'ai')
  eq('块数只剩 1', join.blocks.length, 1)

  // 反向选区（从后往前选）
  const back = docFrom('abc', 'def', 'ghi')
  eq('反向选区也认', JSON.stringify(deleteSpan(back, { blockId: 't2', offset: 2 }, { blockId: 't0', offset: 1 })), '{"blockId":"t0","offset":1}')
  eq('反向选区结果一致', textOf(back).join('|'), 'ai')

  // 端点越界先夹住
  const clamp = docFrom('abc', 'def')
  deleteSpan(clamp, { blockId: 't0', offset: 99 }, { blockId: 't1', offset: 99 })
  eq('端点越界夹到段尾', textOf(clamp).join('|'), 'abc')

  // 中间夹着分页符 / 分节符；分节符被删要连带 sections 一起走
  const withBreak = docFrom('甲', '乙')
  withBreak.blocks.splice(1, 0, { t: 'pageBreak', id: 'pg1' })
  withBreak.blocks.push({ t: 'sectionBreak', id: 's1' }, { t: 'textBlock', id: 't2', kind: 'body', inlines: [{ t: 'text', text: '丙' }] })
  withBreak.sections = [{ orientation: 'landscape' }, {}]
  eq('删前节数对得上', sectionCountOf(withBreak), 2)
  deleteSpan(withBreak, { blockId: 't0', offset: 0 }, { blockId: 't2', offset: 0 })
  eq('换页标记与分节符一并删掉', withBreak.blocks.length, 1)
  eq('sections 跟着分节符一起摘（长度 = 分节符数 + 1）', withBreak.sections?.length, 1)
  eq('留住的是首节的设置', withBreak.sections?.[0]?.orientation, 'landscape')

  // 修订模式：不并段，只有文字被标成 del
  const tracked = docFrom('abc', 'def', 'ghi')
  deleteSpan(tracked, { blockId: 't0', offset: 1 }, { blockId: 't2', offset: 2 }, { kind: 'del', id: 1, author: '甲', date: '' })
  eq('修订模式不并段', tracked.blocks.length, 3)
  eq('被覆盖的文字都留着', textOf(tracked).join('|'), 'abc|def|ghi')
  const marked = (model) =>
    model.blocks
      .filter((b) => b.t === 'textBlock')
      .flatMap((b) => b.inlines)
      .filter((i) => i.t === 'text' && i.rev)
      .map((i) => `${i.text}:${i.rev.kind}`)
      .join(',')
  eq('覆盖到的文字都标成 del', marked(tracked), 'bc:del,def:del,gh:del')

  // 端点不是段落（表格格子）时不认
  const table = parseMd('甲\n\n:::table\n| a | b |\n| c | d |\n:::')
  eq('端点落在格子里 → 不并段', deleteSpan(table, { blockId: 't0', offset: 0 }, { blockId: cellId(table.blocks[1].id, 1, 1), offset: 0 }), null)

  // 中间覆盖到整张表时表格一并删掉（Word 也是这么做的）
  const withTable = parseMd('甲\n\n:::table\n| a | b |\n| c | d |\n:::\n\n乙')
  deleteSpan(withTable, { blockId: withTable.blocks[0].id, offset: 0 }, { blockId: withTable.blocks[2].id, offset: 1 })
  eq('整段覆盖到的表格一起删', withTable.blocks.length, 1)
  eq('两头接起来', plainText(withTable.blocks[0]), '乙'.slice(1))

  // joinWithNext：Delete 落在段尾 = 删掉段落标记
  const next = docFrom('ab', 'cd')
  eq('并下一段的落点', JSON.stringify(joinWithNext(next, 't0')), '{"blockId":"t0","offset":2}')
  eq('并下一段的结果', textOf(next).join('|'), 'abcd')

  // 下一块不是段落（表格 / 换页标记 / 文末）时不接管
  const last = docFrom('ab')
  eq('文末没有下一段', joinWithNext(last, 't0'), null)
  const withPg = docFrom('ab', 'cd')
  withPg.blocks.splice(1, 0, { t: 'pageBreak', id: 'pg2' })
  eq('下一块是换页标记', joinWithNext(withPg, 't0'), null)
  const withTable2 = parseMd('甲\n\n:::table\n| a | b |\n| c | d |\n:::')
  eq('下一块是表格', joinWithNext(withTable2, withTable2.blocks[0].id), null)

  /*
   * 两个谓词：调用方必须在**改模型之前**问它们 —— 撤销快照要记在改模型之前，
   * 而 mergeIntoPrevious / joinWithNext 一进去就动模型（问都不问就记，撤销会变成空操作）。
   */
  const can = docFrom('ab', 'cd')
  eq('前一块是段落 → 可以并', canMergeIntoPrevious(can, 't1'), true)
  eq('首块没有前一块 → 不能并', canMergeIntoPrevious(can, 't0'), false)
  eq('段尾有下一段 → 可以并', canJoinWithNext(can, 't0'), true)
  eq('末块没有下一段 → 不能并', canJoinWithNext(can, 't1'), false)
  const canBreak = docFrom('ab', 'cd')
  canBreak.blocks.splice(1, 0, { t: 'pageBreak', id: 'pgx' })
  eq('前一块是换页标记 → 不能并', canMergeIntoPrevious(canBreak, 't1'), false)
  eq('下一块是换页标记 → 不能并', canJoinWithNext(canBreak, 't0'), false)
  const canTable = parseMd('甲\n\n:::table\n| a | b |\n| c | d |\n:::\n\n乙')
  const tableAt = canTable.blocks.findIndex((b) => b.t === 'table')
  eq('下一块是表格 → 不能并', canJoinWithNext(canTable, canTable.blocks[0].id), false)
  eq('前一块是表格 → 不能并', canMergeIntoPrevious(canTable, canTable.blocks[tableAt + 1].id), false)
  eq('不存在的块 → 两个谓词都说不', `${canMergeIntoPrevious(canTable, 'nope')}/${canJoinWithNext(canTable, 'nope')}`, 'false/false')
}

console.log('\n=== 26. ::editor：编辑器开关的解析与序列化（W6）===')
{
  const shapeOf = (doc) =>
    JSON.stringify({
      blocks: normalizeBlocks(doc),
      sections: doc.sections ?? null,
      editor: doc.editor ?? null,
    })

  // 默认值不写：文档里没有这一行时模型里也没有 editor
  const plain = parseMd('甲')
  eq('没有 ::editor 就没有 editor 字段', plain.editor, undefined)
  eq('默认也不写回 md', toMd(plain), '甲')
  eq('默认开关解析成「修订关、导航开」', JSON.stringify(resolveEditorFlags(plain.editor)), JSON.stringify({ trackChanges: false, nav: true }))

  // 只写非默认值：trackChanges=on 单独出现
  const track = parseMd('::editor trackChanges=on\n\n甲')
  eq('trackChanges=on 读出来', JSON.stringify(track.editor), '{"trackChanges":true}')
  eq('序列化写回一行（在最前面）', toMd(track), '::editor trackChanges=on\n甲')
  eq('再往返一次字节稳定', toMd(parseMd(toMd(track))), '::editor trackChanges=on\n甲')
  eq('往返结构一致', shapeOf(parseMd(toMd(track))), shapeOf(track))

  // 只写非默认值：nav=off 单独出现
  const nav = parseMd('::editor nav=off\n\n甲')
  eq('nav=off 读出来', JSON.stringify(nav.editor), '{"nav":false}')
  eq('序列化只写 nav=off', toMd(nav), '::editor nav=off\n甲')
  eq('nav=off 的开关解析', JSON.stringify(resolveEditorFlags(nav.editor)), JSON.stringify({ trackChanges: false, nav: false }))

  // 两者同时出现，且 key 顺序固定 trackChanges → nav
  const both = parseMd('::editor nav=off trackChanges=on\n\n甲')
  eq('两个键都读出来', JSON.stringify(both.editor), '{"trackChanges":true,"nav":false}')
  eq('序列化顺序固定 trackChanges 在前', toMd(both), '::editor trackChanges=on nav=off\n甲')

  // 显式写出的默认值被归一化掉（与围栏 minLines=1 / 格内 {@listItem|…} 同一个约定）
  const obvious = parseMd('::editor trackChanges=off nav=on\n\n甲')
  eq('显式写出的默认值不落字段', obvious.editor, undefined)
  eq('也不写回 md', toMd(obvious), '甲')

  // 未知 key / 值忽略，不报错也不落字段
  const junk = parseMd('::editor nope=1 trackChanges=maybe nav=off\n\n甲')
  eq('未知 key 被忽略、值只认 on/off', JSON.stringify(junk.editor), '{"nav":false}')

  // 非首行出现的 ::editor 落回正文（与「围栏没闭合绝不当表格」同一条原则）
  const body = parseMd('甲\n\n::editor trackChanges=on')
  eq('正文里的 ::editor 不当指令', toMd(body), '甲\n::editor trackChanges=on')
  eq('也不产生 editor 字段', body.editor, undefined)
  // 指令区里两行都在最前面，谁先谁后都认
  const withSection = parseMd('::editor nav=off\n\n::section numbers=off\n\n甲')
  eq('::editor 与 ::section 同住指令区（nav）', JSON.stringify(withSection.editor), '{"nav":false}')
  eq('::editor 与 ::section 同住指令区（section）', JSON.stringify(settingsOf(withSection, 0)), '{"pageNumbers":false}')
  eq('序列化把 ::editor 写在 ::section 之前', toMd(withSection), '::editor nav=off\n::section numbers=off\n甲')
  eq('交叉顺序往返也稳定', shapeOf(parseMd('::section numbers=off\n::editor nav=off\n\n甲')), shapeOf(withSection))

  // 正文首行之后的 ::editor 不再被认（指令区已关）
  const afterContent = parseMd('::editor nav=off\n\n甲\n\n::editor trackChanges=on')
  eq('指令区之后的 ::editor 落回正文', afterContent.editor ? JSON.stringify(afterContent.editor) : '(无)', '{"nav":false}')
  eq('正文那一行留在块里', toMd(afterContent), '::editor nav=off\n甲\n::editor trackChanges=on')

  // cloneDoc 必须连 editor 一起拷：撤销栈 / 渲染快照都走它
  const cloned = cloneDoc(both)
  eq('cloneDoc 拷了 editor', JSON.stringify(cloned.editor), JSON.stringify(both.editor))
  if (cloned.editor) cloned.editor.nav = true
  eq('拷贝之后改它不影响原稿', JSON.stringify(both.editor), '{"trackChanges":true,"nav":false}')

  /*
   * `::editor` 只影响编辑器界面，**不许影响导出的 docx**。比内层 XML 而不是整个 zip：
   * zip 里带 docx 库写的创建时间戳，两次导出的字节本来就不会相同。
   */
  const spec = resolveSpec()
  const plainDoc = parseMd('# 甲\n\n乙')
  const flagDoc = parseMd('::editor trackChanges=on nav=off\n\n# 甲\n\n乙')
  eq('两种形状的块数一致', flagDoc.blocks.length, plainDoc.blocks.length)
  const docxXml = async (doc) => {
    const base64 = await toBase64(doc, spec, { title: '开关不进 docx' })
    const zip = await JSZip.loadAsync(Buffer.from(base64, 'base64'))
    return (
      (await zip.file('word/document.xml')?.async('string')) +
      '\u0000' +
      (await zip.file('word/styles.xml')?.async('string'))
    )
  }
  const plainXml = await docxXml(plainDoc)
  const flagXml = await docxXml(flagDoc)
  ok(
    '带 ::editor 的 docx 内层 XML 与不带的一模一样',
    plainXml === flagXml,
    `长度 ${plainXml.length} vs ${flagXml.length}`,
  )
}

console.log('\n=== 27. 快捷键表：解析 / 匹配 / 覆盖 / 冲突（W6）===')
{
  const ev = (key, opts = {}) => ({
    key,
    code: opts.code ?? '',
    ctrlKey: opts.ctrl ?? false,
    metaKey: opts.meta ?? false,
    shiftKey: opts.shift ?? false,
    altKey: opts.alt ?? false,
  })

  // ---- parseCombo：大小写不敏感、修饰键归一、主键 ----
  eq('Ctrl+B', JSON.stringify(parseCombo('Ctrl+B')), '{"mod":true,"shift":false,"alt":false,"key":"b"}')
  eq('大小写不敏感', JSON.stringify(parseCombo('cTrL+b')), JSON.stringify(parseCombo('Ctrl+B')))
  eq('Cmd / Meta / ⌘ / Win 都归一到 mod', [
    parseCombo('Cmd+U'),
    parseCombo('Meta+U'),
    parseCombo('⌘+U'),
    parseCombo('Win+U'),
    parseCombo('Ctrl+U'),
  ].every((c) => c && c.mod && c.key === 'u'), true)
  eq('Alt+4 的数字键', JSON.stringify(parseCombo('Alt+4')), '{"mod":false,"shift":false,"alt":true,"key":"4"}')
  eq('F4 单独一个键', JSON.stringify(parseCombo('F4')), '{"mod":false,"shift":false,"alt":false,"key":"f4"}')
  eq('前后空格无所谓', JSON.stringify(parseCombo(' Ctrl + Shift + E ')), '{"mod":true,"shift":true,"alt":false,"key":"e"}')
  eq('空串认不出来', parseCombo(''), null)
  eq('只有修饰键认不出来', parseCombo('Ctrl+'), null)
  eq('认不出的修饰键不猜（Foo+B 不是加粗）', parseCombo('Foo+B'), null)
  eq('主键不合法认不出来', parseCombo('Ctrl+Shift'), null)

  // ---- matchShortcut：修饰键逐项严格比对 ----
  const defaultBold = parseCombo(DEFAULT_SHORTCUTS.bold)
  eq('ctrl+B 命中', matchShortcut(ev('b', { ctrl: true }), defaultBold), true)
  eq('大写的 B 也命中', matchShortcut(ev('B', { ctrl: true }), defaultBold), true)
  eq('meta+B（Cmd）也命中', matchShortcut(ev('b', { meta: true }), defaultBold), true)
  eq('多按了 Shift 就不命中', matchShortcut(ev('B', { ctrl: true, shift: true }), defaultBold), false)
  eq('不带修饰键不命中', matchShortcut(ev('b'), defaultBold), false)
  eq('没绑（null）不命中', matchShortcut(ev('b', { ctrl: true }), null), false)
  const alt4 = parseCombo(DEFAULT_SHORTCUTS.formatAmount)
  eq('Alt+4 的 key 命中', matchShortcut(ev('4', { alt: true }), alt4), true)
  eq('Alt+4 的 code 命中（布局/输入法下 key 不是数字）', matchShortcut(ev('。', { code: 'Digit4', alt: true }), alt4), true)
  eq('code 不是那枚数字就不命中', matchShortcut(ev('。', { code: 'Digit5', alt: true }), alt4), false)
  const f4 = parseCombo(DEFAULT_SHORTCUTS.repeat)
  eq('F4 命中', matchShortcut(ev('F4', { code: 'F4' }), f4), true)

  // ---- 命名主键（W9）：方向键与 Home 用名字或箭头写都认，标签显示成箭头 ----
  const altLeft = parseCombo(DEFAULT_SHORTCUTS.cellAlignTop)
  eq(
    'Ctrl+Alt+← / Left / arrowleft 是同一个',
    [
      parseCombo('Ctrl+Alt+←'),
      parseCombo('Ctrl+Alt+Left'),
      parseCombo('Ctrl+Alt+arrowleft'),
    ]
      .map((combo) => JSON.stringify(combo))
      .join('|'),
    `${JSON.stringify(parseCombo('Ctrl+Alt+Left'))}|${JSON.stringify(parseCombo('Ctrl+Alt+Left'))}|${JSON.stringify(parseCombo('Ctrl+Alt+Left'))}`,
  )
  eq(
    'Ctrl+Alt+← 命中 ArrowLeft 事件',
    matchShortcut(ev('ArrowLeft', { code: 'ArrowLeft', ctrl: true, alt: true }), altLeft),
    true,
  )
  eq('裸 ArrowLeft 不命中（修饰键逐项严格比对）', matchShortcut(ev('ArrowLeft'), altLeft), false)
  eq(
    '标签里方向键显示成箭头',
    comboLabel(parseCombo(DEFAULT_SHORTCUTS.cellAlignMiddle)),
    'Ctrl+Alt+Home',
  )

  // ---- 默认表：29 个动作（W6 的九个 + W9 追加的二十个）都在，且与默认表逐条对齐 ----
  eq('动作集有 29 个', SHORTCUT_ACTIONS.length, 29)
  const table = resolveShortcuts()
  eq(
    '默认表逐条对齐',
    SHORTCUT_ACTIONS.map((a) => `${a}=${table[a] ? DEFAULT_SHORTCUTS[a] : 'null'}`).join(','),
    SHORTCUT_ACTIONS.map((a) => `${a}=${DEFAULT_SHORTCUTS[a]}`).join(','),
  )
  eq('默认表里没有空绑定', SHORTCUT_ACTIONS.every((a) => table[a] !== null), true)

  // ---- 覆盖：改绑生效、旧组合失效 ----
  const warns = []
  const realWarn = console.warn
  console.warn = (...args) => warns.push(args.join(' '))
  try {
    const custom = resolveShortcuts({ bold: 'ctrl+shift+b' })
    eq('新组合生效', matchShortcut(ev('B', { ctrl: true, shift: true }), custom.bold), true)
    eq('旧组合失效', matchShortcut(ev('b', { ctrl: true }), custom.bold), false)
    eq('别的动作没被牵连', matchShortcut(ev('u', { ctrl: true }), custom.underline), true)
    eq('只改了一个动作时会 warn？', warns.length, 0)

    // 未知动作名忽略 + warn，默认表不受影响
    warns.length = 0
    const unknown = resolveShortcuts({ bolld: 'Ctrl+B' })
    eq('未知动作名不悄悄改默认表（ctrl+B 仍加粗）', matchShortcut(ev('b', { ctrl: true }), unknown.bold), true)
    eq('未知动作名 warn 一句', warns.length, 1)
    eq('warn 里点了名', warns[0].includes('bolld'), true)

    // 认不出来的组合键：warn + 保留默认
    warns.length = 0
    const broken = resolveShortcuts({ underline: 'Ctrl+Shift' })
    eq('认不出来的组合键保留默认', matchShortcut(ev('u', { ctrl: true }), broken.underline), true)
    eq('认不出来要 warn', warns.length, 1)

    // 冲突：先到先得（动作顺序 bold 在 underline 之前）
    warns.length = 0
    const clash = resolveShortcuts({ bold: 'Ctrl+U' })
    eq('先到的保住这个键', matchShortcut(ev('u', { ctrl: true }), clash.bold), true)
    eq('后到的解绑（不是偷偷退回默认）', clash.underline, null)
    eq('冲突 warn 一句', warns.length, 1)

    // 组合键字符串本身不合法但能解析的边界：F 键与单字符
    eq('F9 认得出', JSON.stringify(parseCombo('f9')), '{"mod":false,"shift":false,"alt":false,"key":"f9"}')
    eq('F25 不存在', parseCombo('F25'), null)
  } finally {
    console.warn = realWarn
  }
}

console.log('\n=== 28. 表格复选：矩形块 / 目标格列表 / 批量只改选中的格（W7）===')
{
  // makeTable / bodyRow / roleRow 用文件开头那套（格内多段落形状）
  const keys = (cells) => cells.map((c) => `${c.row},${c.col}`).join('|')
  const grid = makeTable(
    [roleRow('unit', 'u'), bodyRow('a', 'b', 'c'), bodyRow('d', 'e', 'f'), roleRow('note', 'n')],
    3,
  )

  // ---- 归一化：unit / note 行整行一格，列归 0 ----
  eq('body 行的列照原样', normalizeCellCol(grid, 1, 2), 2)
  eq('unit 行的列归 0', normalizeCellCol(grid, 0, 2), 0)
  eq('note 行的列归 0', normalizeCellCol(grid, 3, 1), 0)
  eq('越界行不炸（按非 body 处理）', normalizeCellCol(grid, 9, 1), 0)

  // ---- 矩形：端点顺序无关 ----
  eq(
    '两端点上下颠倒也是同一个矩形',
    JSON.stringify(cellRectBetween({ row: 2, col: 2 }, { row: 1, col: 0 })),
    JSON.stringify({ r1: 1, c1: 0, r2: 2, c2: 2 }),
  )

  // ---- 2×2：刷选判据（浏览器侧那条断言的 node 侧对照） ----
  const twoByTwo = cellsInRect(grid, cellRectBetween({ row: 1, col: 1 }, { row: 2, col: 2 }))
  eq('2×2 矩形 = 4 格', twoByTwo.length, 4)
  eq('2×2 矩形按行优先展开', keys(twoByTwo), '1,1|1,2|2,1|2,2')
  eq('矩形自己怎么给顺序都一样', keys(cellsInRect(grid, { r1: 2, c1: 2, r2: 1, c2: 1 })), '1,1|1,2|2,1|2,2')

  // ---- plain 行（unit / note）只有第 0 格：列区间要覆盖第 0 列才算选中它 ----
  eq(
    '列区间含 0 时 unit 行被算进来',
    keys(cellsInRect(grid, { r1: 0, c1: 0, r2: 1, c2: 2 })),
    '0,0|1,0|1,1|1,2',
  )
  eq(
    '列区间不含 0 时不选中 unit 行',
    keys(cellsInRect(grid, { r1: 0, c1: 1, r2: 1, c2: 2 })),
    '1,1|1,2',
  )
  eq('note 行同理（整行算一格）', keys(cellsInRect(grid, { r1: 3, c1: 0, r2: 3, c2: 2 })), '3,0')

  // ---- 参差行：只数模型里真有的格（渲染补出来的幻影格不算） ----
  const ragged = makeTable([bodyRow('a', 'b', 'c'), bodyRow('d')], 3)
  eq('参差行不数幻影格', keys(cellsInRect(ragged, { r1: 0, c1: 0, r2: 1, c2: 2 })), '0,0|0,1|0,2|1,0')
  eq('整块越界 = 空集合', cellsInRect(ragged, { r1: 5, c1: 0, r2: 9, c2: 2 }).length, 0)

  // ---- 多块：去重 + 行优先 ----
  const blocks = [
    { r1: 1, c1: 0, r2: 1, c2: 1 },
    { r1: 2, c1: 1, r2: 2, c2: 2 },
    { r1: 1, c1: 1, r2: 2, c2: 1 },
  ]
  eq('多块合成选中集合（去重 + 行优先）', keys(cellsInRects(grid, blocks)), '1,0|1,1|2,1|2,2')
  eq(
    '同一格出现在两块里只算一次',
    cellsInRects(grid, [
      { r1: 1, c1: 1, r2: 1, c2: 1 },
      { r1: 1, c1: 1, r2: 1, c2: 1 },
    ]).length,
    1,
  )
  eq('空块列表 = 空集合', cellsInRects(grid, []).length, 0)
  eq('sortCells 只排序去重、不改原数组', keys(sortCells([{ row: 2, col: 1 }, { row: 1, col: 2 }])), '1,2|2,1')

  // ---- cellRectIndexOf：Ctrl+点击「已在选中的格」= 去掉包含它的那一块 ----
  eq('找得到包含某格的块', cellRectIndexOf(grid, blocks, { row: 2, col: 2 }), 1)
  eq('第一块也能找回来', cellRectIndexOf(grid, blocks, { row: 1, col: 0 }), 0)
  eq('没被任何块包含就是 -1', cellRectIndexOf(grid, blocks, { row: 3, col: 0 }), -1)

  // ---- 批量落笔：只改选中的格，隔壁一个都不许动 ----
  const target = cellsInRect(grid, cellRectBetween({ row: 1, col: 1 }, { row: 2, col: 2 }))
  const untouched = JSON.stringify({ unit: grid.rows[0], note: grid.rows[3], col0: [grid.rows[1].cells[0], grid.rows[2].cells[0]] })
  setCellsKind(grid, target, 'h2')
  eq(
    '批量换样式只改选中的 4 格',
    grid.rows
      .flatMap((row, ri) => row.cells.map((cell, ci) => (cell.kind ? `${ri},${ci}` : '')))
      .filter((s) => s !== '')
      .join('|'),
    '1,1|1,2|2,1|2,2',
  )
  eq(
    '隔壁格（含 unit / note 行）一个都没动',
    JSON.stringify({ unit: grid.rows[0], note: grid.rows[3], col0: [grid.rows[1].cells[0], grid.rows[2].cells[0]] }),
    untouched,
  )
  eq('没改的格不落冗余字段', grid.rows[1].cells[0].kind, undefined)

  setCellsAlign(grid, target, 'h', 'center')
  eq('批量水平对齐写进选中的 4 格', target.every(({ row, col }) => grid.rows[row].cells[col].align?.h === 'center'), true)
  eq('没选中的格没有 align 字段', grid.rows[1].cells[0].align, undefined)
  eq('批量垂直对齐另一维', (setCellsAlign(grid, target, 'v', 'middle'), target.every(({ row, col }) => grid.rows[row].cells[col].align?.v === 'middle')), true)
  setCellsAlign(grid, target, 'h', null)
  setCellsAlign(grid, target, 'v', null)
  eq(
    '清除覆盖（null）两维都没了就连 align 字段一起删',
    target.every(({ row, col }) => grid.rows[row].cells[col].align === undefined),
    true,
  )

  // ---- 预判：空转就不该记撤销（组件据此返回 false） ----
  eq('本来就没覆盖 → 没有要改的格', cellsChangingAlign(grid, target, 'h', null).length, 0)
  eq('要写一个不同的值 → 4 格都要改', cellsChangingAlign(grid, target, 'h', 'left').length, 4)
  eq('样式已经一致 → 没有要改的格', cellsChangingKind(grid, target, 'h2').length, 0)
  eq('样式不一致 → 4 格都要改', cellsChangingKind(grid, target, 'body').length, 4)
  eq('预判也按行优先', keys(cellsChangingAlign(grid, target, 'h', 'left')), '1,1|1,2|2,1|2,2')

  // ---- 批量对齐的「再点同一个值 = 清除覆盖」规则 ----
  eq('整批都等于目标值 → 清除（null）', nextAlignValue(['center', 'center'], 'center'), null)
  eq('有一格不同 → 一律写目标值', nextAlignValue(['center', 'left'], 'center'), 'center')
  eq('空批不算「一致」（不能误清）', nextAlignValue([], 'center'), 'center')

  // ---- 存储值（预判与落笔共用同一个读法；只读 kind / align，与格内段落无关） ----
  eq('storedCellKind 的缺省语义是 listItem', storedCellKind({}), 'listItem')
  eq('storedCellAlign 缺省 null', storedCellAlign({}, 'h'), null)
  eq('storedCellAlign 读得到覆盖', storedCellAlign({ align: { v: 'bottom' } }, 'v'), 'bottom')
}

console.log('\n=== 29. 表头 / 附注行恒「最小一行」：CSS 侧的规则（W7 第②条）===')
{
  // 这条规则的另一半在 docx/export.ts（w:trHeight），由 verify:docx / verify:p1 在字节与磅值层核对；
  // 这里盯的是预览 CSS 的生成结果 —— 两条规则谁盖谁，是纯字符串就能定下来的事。
  const css = buildCss(resolveSpec())
  const lines = css.split('\n')
  const min2Rules = lines.filter((line) => line.includes('wtp-table-min2 td.'))
  eq('-min2 的行高规则逐条样式生成（条数 = 样式数）', min2Rules.length, STYLE_KEYS.length)
  ok(
    '每条 -min2 规则都排除了 plain 行（否则表头/附注行会被撑成两行）',
    min2Rules.every((line) => line.includes(':not(.wtp-td-plain)')),
    min2Rules.join(' / '),
  )
  eq(
    'plain 行的下限仍由不带 -min2 的那条规则提供',
    lines.filter((line) => /^\.wtp-table td\.wtp-td-\w+ \{ height: [\d.]+pt; \}$/.test(line)).length,
    STYLE_KEYS.length,
  )
  ok(
    '整格复选的高亮只给底色，不碰 padding / height / border',
    lines.some(
      (line) =>
        new RegExp(`^\\.wtp-table td\\.${CELL_SELECTION_CLASS} \\{ background: [^}]+\\}$`).test(
          line,
        ),
    ),
    lines.find((line) => line.includes(CELL_SELECTION_CLASS)) ?? '(没有这条规则)',
  )
}

console.log('\n=== 30. 默认快捷键表落在一个可直接手改的 json 上 ===')
{
  // 默认表的唯一真相源是 src/lib/edit/shortcuts.json（它是 `shortcuts` prop 的默认值）。
  // 这里盯四件事：json 与代码里的常量一致、json 里每个值都解析得出来、
  // 键集合与动作集严丝合缝，以及「整份 json 当 prop 传」与「不传」完全等价 ——
  // 否则「json 就是 prop 的默认值」这句话在语义上就是不成立的。
  const jsonUrl = new URL('../src/lib/edit/shortcuts.json', import.meta.url)
  const raw = JSON.parse(readFileSync(jsonUrl, 'utf8'))
  eq(
    'json 的键集合 = 动作集（无遗漏、无多余）',
    Object.keys(raw).sort().join(','),
    [...SHORTCUT_ACTIONS].sort().join(','),
  )
  ok(
    'DEFAULT_SHORTCUTS 逐键等于 json 里的值',
    SHORTCUT_ACTIONS.every((action) => DEFAULT_SHORTCUTS[action] === raw[action]),
    JSON.stringify(raw),
  )
  ok(
    'json 里的每个组合键都解析得出来（手改出坏值不会静默失效）',
    SHORTCUT_ACTIONS.every((action) => parseCombo(raw[action]) !== null),
    JSON.stringify(raw),
  )
  eq(
    '整份 json 当 prop 传 = 不传（默认表就是它）',
    JSON.stringify(resolveShortcuts(raw)),
    JSON.stringify(resolveShortcuts()),
  )
  ok(
    '默认表里 29 个动作都绑上了键',
    SHORTCUT_ACTIONS.every((action) => resolveShortcuts()[action] !== null),
    SHORTCUT_ACTIONS.filter((action) => resolveShortcuts()[action] === null).join(','),
  )
  /*
   * W9 的护栏：手动往 json 里排 29 个键位，最容易犯的错就是两个动作写成同一个组合键 ——
   * resolveShortcuts 会按「先到先得」把后到的解绑（默认表就静默少一个键），所以这里
   * 把两件事一起钉住：每个动作都绑上了键，且 29 个组合键的标签两两不同。
   */
  const defaultLabels = SHORTCUT_ACTIONS.map((action) => {
    const combo = resolveShortcuts()[action]
    return combo ? comboLabel(combo) : `未绑定:${action}`
  })
  eq(
    '默认表 29 个动作两两不冲突（都有键、标签互不相同）',
    `${defaultLabels.filter((label) => !label.startsWith('未绑定')).length} 个有键 / ${new Set(defaultLabels).size} 个不同标签`,
    '29 个有键 / 29 个不同标签',
  )
}

console.log(`\n通过 ${passed} 项，失败 ${failed} 项`)
process.exit(failed === 0 ? 0 : 1)


