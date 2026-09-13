/**
 * 编辑操作模型层的单元测试。
 *
 * lib/edit/model.ts 是纯函数（只吃模型、不改 DOM），所以能在 node 里直接验。
 * 这里盯的是几条容易悄悄出错的地方：
 *   · 按区间替换时，夹住区间的批注锚点必须成对留下（否则渲染会把余下的文字吞进高亮）；
 *   · 切分 / 合并的边界（开头、结尾、加粗/下划线等内联格式的边界）；
 *   · 加粗、下划线、颜色的判断与增删；
 *   · 金额格式化（千分位 + 两位小数）的取舍；
 *   · 修订模式下删除不真删，而是标成 del；
 *   · 特殊空格（U+2003/2002/2005）能原样写进 docx 的 document.xml。
 *
 * 用法：node scripts/test-edit-model.mjs   （需先 npm run build:lib）
 */

import JSZip from 'jszip'

import {
  addComment,
  applyFormat,
  blockLength,
  bodyInsertIndex,
  bodyRowIndexes,
  buildOutline,
  cellId,
  cloneDoc,
  commentScopes,
  containerLength,
  deleteRange,
  findBlock,
  findCell,
  findContainer,
  findMatches,
  findTable,
  formatAmount,
  insertBodyRow,
  insertBreakAfter,
  insertColumn,
  insertText,
  mergeIntoPrevious,
  normalizeBlocks,
  normalizeTable,
  outlineSignature,
  parseCellId,
  parseMd,
  plainText,
  rangeColor,
  rangeIsBold,
  rangeIsUnderline,
  removeBodyRow,
  removeBreak,
  removeColumn,
  removeComment,
  removeTable,
  renderTableFragment,
  replyComment,
  replaceMatches,
  replaceRange,
  resolveSpec,
  setBlockKind,
  setCellAlign,
  setCellKind,
  setContainerKind,
  setMinLines,
  setRoleRow,
  sliceInlines,
  splitBlock,
  stepCell,
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
  ok('分节符默认重排页码', last.t === 'sectionBreak' && last.restartNumbering === true)

  eq('分页符删得掉', removeBreak(model, pageId), true)
  eq('分页符没了', model.blocks.some((b) => b.t === 'pageBreak'), false)
  eq('文字块不归 removeBreak 管', removeBreak(model, 't0'), false)

  // 附件与换页标记都要能过 md 往返
  const round = {
    blocks: [
      { t: 'textBlock', id: 'a', kind: 'attachment', inlines: [{ t: 'text', text: '附件一' }] },
      { t: 'pageBreak', id: 'p' },
      { t: 'sectionBreak', id: 's', restartNumbering: true },
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
  const cellText = (cell) =>
    cell.inlines.filter((i) => i.t === 'text').map((i) => i.text).join('')

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
  eq('列标题加粗', table.rows[1].cells[0].inlines[0].bold, true)
  eq('格里的竖线没被当分隔符', cellText(table.rows[3].cells[0]), '备注|说明')
  eq('格里的反斜杠原样还原', cellText(table.rows[3].cells[1]), '含\\反斜杠')

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

  eq('cellId', cellId('tb1', 2, 3), 'tb1.r2c3')
  eq(
    'parseCellId 往返',
    JSON.stringify(parseCellId('tb1.r2c3')),
    JSON.stringify({ tableId: 'tb1', row: 2, col: 3 }),
  )
  eq('parseCellId 拒绝非单元格 id', parseCellId('b1'), null)

  // cloneDoc 必须逐层新建：撤销栈与渲染快照都靠它，共享引用会被后续编辑改到
  const copy = cloneDoc(model)
  copy.blocks[0].rows[1].cells[0].inlines[0].text = '改过了'
  copy.blocks[0].rows.push({ role: 'body', cells: [{ inlines: [] }] })
  eq('副本的行数变了', copy.blocks[0].rows.length, 6)
  eq('原件行数没变', model.blocks[0].rows.length, 5)
  eq('原件格文字没变', cellText(model.blocks[0].rows[1].cells[0]), '项目')
  eq('不共享 rows 数组', copy.blocks[0].rows === model.blocks[0].rows, false)
  eq('不共享 cells 数组', copy.blocks[0].rows[1].cells === model.blocks[0].rows[1].cells, false)
  eq(
    '不共享 inlines 数组',
    copy.blocks[0].rows[1].cells[0].inlines === model.blocks[0].rows[1].cells[0].inlines,
    false,
  )
}

console.log('\n=== 17b. 表格：格内行内语法 / 首尾空格 / 未闭合围栏 ===')
{
  // 这些用例是独立验收方补的：初版 splitTableCells 只看反斜杠，不跳 {} 与 [[]]，
  // 于是格内 {红|甲}、[[甲|核对原件]] 里的竖线被当成列分隔符 —— 一格拆多格、颜色与批注静默丢失。
  const cellText = (cell) =>
    cell.inlines.filter((i) => i.t === 'text').map((i) => i.text).join('')
  const asDoc = (block) => ({ blocks: [block], comments: [] })
  const sameShape = (a, b) =>
    JSON.stringify(normalizeBlocks(a)) === JSON.stringify(normalizeBlocks(b))

  const colored = parseMd(':::table\n| {红|甲} | b |\n:::')
  eq('彩色格只算一格', colored.blocks[0].rows[0].cells.length, 2)
  eq('彩色格列数', colored.blocks[0].columns, 2)
  eq('彩色格文字', cellText(colored.blocks[0].rows[0].cells[0]), '甲')
  eq('彩色格颜色读出来', colored.blocks[0].rows[0].cells[0].inlines[0].color, 'FF0000')
  eq('彩色格往返一致', sameShape(parseMd(toMd(colored)), colored), true)

  const hex = parseMd(':::table\n| {#00FF00|乙} | c |\n:::')
  eq('十六进制色号格也只算一格', hex.blocks[0].rows[0].cells.length, 2)
  eq('十六进制色号读出来', hex.blocks[0].rows[0].cells[0].inlines[0].color, '00FF00')

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
  eq('未闭合围栏的后文还在', cellText(unterminated.blocks[2]), '普通一段')
}

console.log('\n=== 17c. md 往返：批注内容里的 \\ 与 | 必须还原（不是越滚越多）===')
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
}

console.log('\n=== 22. 表格片段渲染：接口约束（外层无 data-block-id、格内坐标、行区间）===')
{
  // 这些约束是「量测与预览共用一套 DOM」的地基：外层若挂了 data-block-id，
  // fragmentOf 会把整张表当成一个片段读回模型；格内若不挂，格内根本编辑不了。
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
  eq('格内都挂了 data-block-id', (html.match(/data-block-id="/g) ?? []).length, 2 + 1 + 2 + 1)
  eq(
    '格内坐标是 cellId 形态且 from=0',
    html.includes(`data-block-id="${cellId(table.id, 1, 0)}" data-from="0" data-to="1"`),
    true,
  )
  eq('软换行渲染成带类的 <br>（与空段落占位区分）', html.includes('<br class="wtp-br">'), true)

  const partial = renderTableFragment(table, 1, 3)
  eq('按行区间只渲那两行', (partial.match(/<tr/g) ?? []).length, 2)
  eq('行区间从 1 开始时格内 id 跟着行号走', partial.includes(`data-block-id="${cellId(table.id, 1, 0)}"`), true)

  const single = parseMd(':::table\n| a |\n:::').blocks[0]
  eq('minLines=1 用另一个修饰类', renderTableFragment(single, 0, 1).includes('wtp-table-min1'), true)
}

console.log('\n=== 23. 表格结构操作：增删行/列、unit&note 开关、归一化、行高 ===')
{
  /** 直接造一张表（不走 md，才能造出「各 body 行格数参差」这类形状） */
  const makeTable = (rows, columns = 1, minLines = 1) => ({
    t: 'table',
    id: 'tb1',
    rows,
    columns,
    minLines,
    cantSplit: true,
  })
  const bodyRow = (...texts) => ({
    role: 'body',
    cells: texts.map((text) => ({ inlines: text === '' ? [] : [{ t: 'text', text }] })),
  })
  const roleRow = (role, text = '') => ({
    role,
    cells: [{ inlines: text === '' ? [] : [{ t: 'text', text }] }],
  })
  const shape = (table) => table.rows.map((r) => `${r.role}:${r.cells.length}`).join(',')

  // ---- insertBodyRow ----
  {
    const table = makeTable([bodyRow('a', 'b'), bodyRow('c', 'd'), bodyRow('e', 'f')], 2)
    eq('插在中间返回新行下标', insertBodyRow(table, 1), 1)
    eq('插在中间后行数 +1', table.rows.length, 4)
    eq('新行是 body', table.rows[1].role, 'body')
    eq('新行格数 = 当时的 columns', table.rows[1].cells.length, 2)
    eq('其余行没被挪动', table.rows[2].cells[0].inlines[0].text, 'c')

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
    eq('新列内容为空', table.rows[1].cells[1].inlines.length, 0)
    eq('原格向后挪', table.rows[1].cells[2].inlines[0].text, 'b')

    // 参差行：at 按每行实际格数夹取，不要求先拍平
    const ragged = makeTable([bodyRow('a', 'b'), bodyRow('c', 'd', 'e')], 3)
    insertColumn(ragged, 2)
    eq('参差行各按自己的长度夹取', shape(ragged), 'body:3,body:4')
    eq('columns 重算为最大格数', ragged.columns, 4)

    // removeColumn：只动 body；unit/note 不变
    const del = makeTable([roleRow('unit', 'u'), bodyRow('a', 'b', 'c'), bodyRow('d', 'e', 'f'), roleRow('note', 'n')], 3)
    eq('删列成功', removeColumn(del, 1), true)
    eq('删列只动 body 行', shape(del), 'unit:1,body:2,body:2,note:1')
    eq('删掉的是第 1 列', del.rows[1].cells.map((c) => c.inlines[0]?.text).join(''), 'ac')
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
    eq('插列后 unit 行的文字还在', zeroCol.rows[0].cells[0].inlines[0].text, 'u')
    eq('插列后 note 行的文字还在', zeroCol.rows[2].cells[0].inlines[0].text, 'n')
    eq('新列插在 body 行的 0 号位', zeroCol.rows[1].cells[0].inlines.length, 0)
    eq('body 行原格向后挪', zeroCol.rows[1].cells[1].inlines[0].text, 'a')

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
    eq('body 行删掉的确实是第 0 列', zeroDel.rows[1].cells[0].inlines[0].text, 'b')
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
    trim.rows[0].cells = [{ inlines: [] }, { inlines: [] }, { inlines: [] }]
    normalizeTable(trim)
    eq('unit 行被裁到只剩第 0 格', trim.rows[0].cells.length, 1)

    const empty = makeTable([bodyRow(), bodyRow('a', 'b')], 0)
    normalizeTable(empty)
    eq('columns 取 body 行的最大值（空格不影响）', empty.columns, 2)
    const zero = makeTable([bodyRow()], 0)
    normalizeTable(zero)
    eq('全是空格时 columns 兜底为 1', zero.columns, 1)
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
  const makeTable = (rows, columns = 1, minLines = 1) => ({
    t: 'table',
    id: 'tb1',
    rows,
    columns,
    minLines,
    cantSplit: true,
  })
  const bodyRow = (...texts) => ({
    role: 'body',
    cells: texts.map((text) => ({ inlines: text === '' ? [] : [{ t: 'text', text }] })),
  })
  const roleRow = (role, text = '') => ({
    role,
    cells: [{ inlines: text === '' ? [] : [{ t: 'text', text }] }],
  })
  const cellText = (cell) =>
    cell.inlines.map((i) => (i.t === 'text' ? i.text : '')).join('')
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
        { t: 'sectionBreak', id: 's1', restartNumbering: true },
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
    const cell = { inlines: [{ t: 'text', text: '甲' }] }
    setCellKind(cell, 'h2')
    eq('setCellKind 设值', cell.kind, 'h2')
    setCellKind(cell, 'listItem')
    eq('回到 listItem 时删掉字段（模型不存冗余值）', 'kind' in cell, false)

    const model = parseMd('正文一段\n\n:::table\n| 甲 | 乙 |\n:::')
    const table = model.blocks.find((b) => b.t === 'table')
    setContainerKind(model, model.blocks[0].id, 'h1')
    eq('setContainerKind 改段落', findBlock(model, model.blocks[0].id).kind, 'h1')
    setContainerKind(model, cellId(table.id, 0, 0), 'h3')
    eq('setContainerKind 改格子', findCell(model, cellId(table.id, 0, 0)).kind, 'h3')

    const html = renderTableFragment(table, 0, table.rows.length)
    eq('格内 div 带稳定钩子 wtp-cell + 该格样式 wtp-h3', html.includes('class="wtp-cell wtp-h3"'), true)
    eq('没改过的格子仍是 wtp-cell wtp-listItem', html.includes('class="wtp-cell wtp-listItem"'), true)

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
    const cell = { inlines: [] }
    setCellAlign(cell, 'h', 'center')
    eq('单维：水平', JSON.stringify(cell.align), JSON.stringify({ h: 'center' }))
    setCellAlign(cell, 'v', 'middle')
    eq('双维', JSON.stringify(cell.align), JSON.stringify({ h: 'center', v: 'middle' }))
    setCellAlign(cell, 'h', null)
    eq('清一维后另一维还在', JSON.stringify(cell.align), JSON.stringify({ v: 'middle' }))
    setCellAlign(cell, 'v', null)
    eq('两维都清掉后 align 字段消失', 'align' in cell, false)
    setCellAlign(cell, 'h', null)
    eq('本来就没有 align 时再清是空操作', 'align' in cell, false)

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

console.log(`\n通过 ${passed} 项，失败 ${failed} 项`)
process.exit(failed === 0 ? 0 : 1)

