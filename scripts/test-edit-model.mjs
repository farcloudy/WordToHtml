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
  buildOutline,
  cellId,
  cloneDoc,
  commentScopes,
  deleteRange,
  findBlock,
  findMatches,
  formatAmount,
  insertBreakAfter,
  insertText,
  mergeIntoPrevious,
  normalizeBlocks,
  outlineSignature,
  parseCellId,
  parseMd,
  plainText,
  rangeColor,
  rangeIsBold,
  rangeIsUnderline,
  removeBreak,
  removeComment,
  replyComment,
  replaceMatches,
  replaceRange,
  resolveSpec,
  setBlockKind,
  splitBlock,
  toBase64,
  toMd,
  updateComment,
  validateQuery,
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

console.log(`\n通过 ${passed} 项，失败 ${failed} 项`)
process.exit(failed === 0 ? 0 : 1)

