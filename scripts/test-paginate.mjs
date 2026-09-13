/**
 * 分页算法的单元测试。
 *
 * paginate 是纯函数（只吃实测值、不碰 DOM），所以能在 node 里直接验；
 * DOM 测量那部分留给浏览器里的人工/端到端检查。
 *
 * 用法：node scripts/test-paginate.mjs
 */

import { paginate } from '../dist-lib/wordtohtml.mjs'

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

/** 造一个测量结果：把 length 个字符均匀分到 rows 行 */
function block(id, { kind = 'body', rows, length, lineHeight = 25, spaceBefore = 0, spaceAfter = 0 }) {
  const rowStarts = []
  for (let r = 0; r < rows; r += 1) rowStarts.push(Math.round((length * r) / rows))
  return {
    t: 'block',
    blockId: id,
    kind,
    displayLength: length,
    rows,
    lineHeight,
    spaceBefore,
    spaceAfter,
    rowStarts,
  }
}

console.log('=== 1. 空文档 ===')
{
  const pages = paginate([], { contentHeight: 100 })
  eq('页数', pages.length, 1)
  eq('页码', pages[0].pageNumber, 1)
  eq('片段数', pages[0].fragments.length, 0)
}

console.log('\n=== 2. 单行段落装箱（100px 版心 / 25px 行高 = 每页 4 行）===')
{
  const items = Array.from({ length: 11 }, (_, i) =>
    block(`b${i}`, { rows: 1, length: 10 }),
  )
  const pages = paginate(items, { contentHeight: 100 })
  eq('页数', pages.length, 3)
  eq('第1页片段数', pages[0].fragments.length, 4)
  eq('第2页片段数', pages[1].fragments.length, 4)
  eq('第3页片段数', pages[2].fragments.length, 3)
  eq('第1页页码', pages[0].pageNumber, 1)
  eq('第2页页码', pages[1].pageNumber, 2)
  eq('第3页页码', pages[2].pageNumber, 3)
}

console.log('\n=== 3. 长段落跨页按行切开 ===')
{
  // 10 行、100 字符；版心 100px 只能放 4 行
  const items = [block('long', { rows: 10, length: 100 })]
  const pages = paginate(items, { contentHeight: 100 })
  eq('页数', pages.length, 3)
  eq('第1页行数', pages[0].fragments[0].to - pages[0].fragments[0].from, 40)
  ok('第1片不是续排', pages[0].fragments[0].continuation === false)
  ok('第2片是续排', pages[1].fragments[0].continuation === true)
  ok('第3片是续排', pages[2].fragments[0].continuation === true)
  eq('第1片起点', pages[0].fragments[0].from, 0)
  eq('第2片起点', pages[1].fragments[0].from, 40)
  eq('第3片起点', pages[2].fragments[0].from, 80)
  eq('末片终点', pages[2].fragments[0].to, 100)
  eq('片段总数', pages.reduce((n, p) => n + p.fragments.length, 0), 3)
}

console.log('\n=== 4. 孤行控制：少于 4 行的段落不拆 ===')
{
  // 第 1 段 3 行占满 75px，第 2 段 3 行只剩 25px（只够 1 行）→ 整段挪到下一页
  const items = [
    block('a', { rows: 3, length: 30 }),
    block('b', { rows: 3, length: 30 }),
  ]
  const pages = paginate(items, { contentHeight: 100 })
  eq('页数', pages.length, 2)
  eq('第2页片段数', pages[1].fragments.length, 1)
  eq('第2页承载的块', pages[1].fragments[0].blockId, 'b')
  eq('第2页未切开', pages[1].fragments[0].continuation, false)
}

console.log('\n=== 5. 孤行控制：4 行以上可以拆，且两侧各留 2 行 ===')
{
  // 版心 4 行；先放 1 行，再放一个 6 行段落 → 剩 3 行够放，但必须给下一页留 2 行 → 本页放 4 行里的…
  const items = [block('head', { rows: 1, length: 10 }), block('six', { rows: 6, length: 60 })]
  const pages = paginate(items, { contentHeight: 100 })
  eq('页数', pages.length, 2)
  // 第 1 页：1 行 head + 3 行 six（6-3=3 ≥ 2 且 3 ≥ 2）
  eq('第1页片段数', pages[0].fragments.length, 2)
  eq('第1页承接的行数', (pages[0].fragments[1].to - pages[0].fragments[1].from) / 10, 3)
  eq('第2页承接的行数', (pages[1].fragments[0].to - pages[1].fragments[0].from) / 10, 3)
}

console.log('\n=== 6. 分节符：新起一页 + 页码重排 ===')
{
  const items = [
    block('a', { rows: 1, length: 10 }),
    { t: 'break', blockId: 's0', kind: 'section', restartNumbering: true },
    block('b', { rows: 1, length: 10 }),
  ]
  const pages = paginate(items, { contentHeight: 100 })
  eq('页数', pages.length, 2)
  eq('第1页节号', pages[0].sectionIndex, 0)
  eq('第2页节号', pages[1].sectionIndex, 1)
  eq('第1页页码', pages[0].pageNumber, 1)
  eq('第2页页码（重排）', pages[1].pageNumber, 1)
  eq('第1页末尾标记为分节符', pages[0].breaks[0]?.kind, 'section')
  eq('标记带着块 id（编辑器要按它删除）', pages[0].breaks[0]?.blockId, 's0')
  eq('第1页只有一枚标记', pages[0].breaks.length, 1)
  eq('第2页没有标记', pages[1].breaks.length, 0)
}

console.log('\n=== 7. 分节符：不重排时页码延续 ===')
{
  const items = [
    block('a', { rows: 1, length: 10 }),
    { t: 'break', blockId: 's1', kind: 'section', restartNumbering: false },
    block('b', { rows: 1, length: 10 }),
  ]
  const pages = paginate(items, { contentHeight: 100 })
  eq('页数', pages.length, 2)
  eq('第2页页码（延续）', pages[1].pageNumber, 2)
}

console.log('\n=== 7.1 分页符：只换页，不新开一节、页码连续 ===')
{
  const items = [
    block('a', { rows: 1, length: 10 }),
    { t: 'break', blockId: 'p0', kind: 'page', restartNumbering: false },
    block('b', { rows: 1, length: 10 }),
  ]
  const pages = paginate(items, { contentHeight: 100 })
  eq('页数', pages.length, 2)
  eq('节号不变', pages[1].sectionIndex, pages[0].sectionIndex)
  eq('第2页页码连续', pages[1].pageNumber, 2)
  eq('第1页末尾标记为分页符', pages[0].breaks[0]?.kind, 'page')
}

console.log('\n=== 7.2 文末分页符：内容不变也要留下标记 ===')
{
  // 只有一段 + 一个文末分页符：片段一模一样，但标记必须出现在这一页末尾。
  // Word 实测：这种情况不会多出一张空白页（1 页），换页落在空页上被吸收。
  const items = [
    block('a', { rows: 1, length: 10 }),
    { t: 'break', blockId: 'p1', kind: 'page', restartNumbering: false },
  ]
  const pages = paginate(items, { contentHeight: 100 })
  eq('仍然只有一页', pages.length, 1)
  eq('末尾标记还在', pages[0].breaks[0]?.blockId, 'p1')
}

console.log('\n=== 7.3 分页符 + 分节符连在一起：两枚标记都要看得见 ===')
{
  // 这一组对应 Word 实测的 pg1：内容 + 分页符 + 分节符 + 内容 → Word 报 2 页 2 节。
  // 两枚标记都落在第 1 页底部（落到空页上的那一枚只吸收换页，不吸收标记）。
  const items = [
    block('a', { rows: 1, length: 10 }),
    { t: 'break', blockId: 'pb', kind: 'page', restartNumbering: false },
    { t: 'break', blockId: 'sc', kind: 'section', restartNumbering: true },
    block('b', { rows: 1, length: 10 }),
  ]
  const pages = paginate(items, { contentHeight: 100 })
  eq('页数（只推进一页，与 Word 一致）', pages.length, 2)
  eq('第1页挂了两枚标记', pages[0].breaks.length, 2)
  eq('第1枚是分页符', pages[0].breaks[0]?.kind, 'page')
  eq('第2枚是分节符', pages[0].breaks[1]?.kind, 'section')
  eq('第2页节号已推进', pages[1].sectionIndex, 1)
  eq('第2页页码重排', pages[1].pageNumber, 1)
  eq('第2页没有标记', pages[1].breaks.length, 0)
}

console.log('\n=== 7.4 分节符在后面 + 分页符在前（顺序反过来）===')
{
  // Word 实测的 pg4：内容 + 分节符 + 分页符 + 内容 → 2 页 2 节，两枚标记都看得见
  const items = [
    block('a', { rows: 1, length: 10 }),
    { t: 'break', blockId: 'sc', kind: 'section', restartNumbering: true },
    { t: 'break', blockId: 'pb', kind: 'page', restartNumbering: false },
    block('b', { rows: 1, length: 10 }),
  ]
  const pages = paginate(items, { contentHeight: 100 })
  eq('页数', pages.length, 2)
  eq('两枚标记都在第1页', pages[0].breaks.length, 2)
  eq('顺序保持插入顺序', pages[0].breaks.map((b) => b.kind).join(','), 'section,page')
}

console.log('\n=== 7.5 文末分节符：Word 会留一张空白页 ===')
{
  // Word 实测的 pg3：内容 + 文末分节符 → 2 页 2 节（第 2 页是空白页）
  const items = [
    block('a', { rows: 1, length: 10 }),
    { t: 'break', blockId: 'sc', kind: 'section', restartNumbering: true },
  ]
  const pages = paginate(items, { contentHeight: 100 })
  eq('页数', pages.length, 2)
  eq('第2页没有片段（空白页）', pages[1].fragments.length, 0)
  eq('第2页属于新节', pages[1].sectionIndex, 1)
  eq('第2页页码重排为 1', pages[1].pageNumber, 1)
  eq('分节符标记在第1页底部', pages[0].breaks[0]?.blockId, 'sc')
  eq('空白页上没有标记', pages[1].breaks.length, 0)
}

console.log('\n=== 7.6 文档一开头就是换页标记：标记挂到第一页 ===')
{
  const items = [
    { t: 'break', blockId: 'pb', kind: 'page', restartNumbering: false },
    block('a', { rows: 1, length: 10 }),
  ]
  const pages = paginate(items, { contentHeight: 100 })
  eq('页数', pages.length, 1)
  eq('标记挂在第一页', pages[0].breaks[0]?.blockId, 'pb')
}

console.log('\n=== 8. 页顶不叠加段前距 ===')
{
  // 段前距 400px 远大于版心 100px：若在页顶也叠加，将无法放置
  const items = [block('a', { rows: 1, length: 10, spaceBefore: 400 })]
  const pages = paginate(items, { contentHeight: 100 })
  eq('页数', pages.length, 1)
  eq('仍然放下', pages[0].fragments.length, 1)
}

console.log('\n=== 9. 行高超过版心时不死循环 ===')
{
  const items = [block('huge', { rows: 1, length: 10, lineHeight: 500 })]
  const pages = paginate(items, { contentHeight: 100 })
  eq('页数', pages.length, 1)
  eq('兜底放下一行', pages[0].fragments.length, 1)
}

console.log(`\n通过 ${passed} 项，失败 ${failed} 项`)
process.exit(failed === 0 ? 0 : 1)
