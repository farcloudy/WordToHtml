/**
 * 编辑操作模型层的单元测试。
 *
 * lib/edit/model.ts 是纯函数（只吃模型、不改 DOM），所以能在 node 里直接验。
 * 这里盯的是几条容易悄悄出错的地方：
 *   · 按区间替换时，夹住区间的批注锚点必须成对留下（否则渲染会把余下的文字吞进高亮）；
 *   · 切分 / 合并的边界（开头、结尾、加粗等内联格式的边界）；
 *   · 加粗的判断、颜色的一致性；
 *   · 修订模式下删除不真删，而是标成 del。
 *
 * 用法：node scripts/test-edit-model.mjs   （需先 npm run build:lib）
 */

import {
  addComment,
  applyFormat,
  blockLength,
  cloneDoc,
  commentScopes,
  deleteRange,
  findBlock,
  insertBreakAfter,
  insertText,
  mergeIntoPrevious,
  normalizeBlocks,
  parseMd,
  plainText,
  rangeColor,
  rangeIsBold,
  removeBreak,
  removeComment,
  replyComment,
  replaceRange,
  setBlockKind,
  splitBlock,
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

console.log(`\n通过 ${passed} 项，失败 ${failed} 项`)
process.exit(failed === 0 ? 0 : 1)
