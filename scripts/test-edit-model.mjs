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
  cloneDoc,
  commentScopes,
  deleteRange,
  findBlock,
  formatAmount,
  insertBreakAfter,
  insertText,
  mergeIntoPrevious,
  normalizeBlocks,
  parseMd,
  plainText,
  rangeColor,
  rangeIsBold,
  rangeIsUnderline,
  removeBreak,
  removeComment,
  replyComment,
  replaceRange,
  resolveSpec,
  setBlockKind,
  splitBlock,
  toBase64,
  toMd,
  updateComment,
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

console.log(`\n通过 ${passed} 项，失败 ${failed} 项`)
process.exit(failed === 0 ? 0 : 1)
