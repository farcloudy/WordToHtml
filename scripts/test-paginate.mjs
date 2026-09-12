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
  const items = [block('a', { rows: 1, length: 10 }), { t: 'break', restartNumbering: true }, block('b', { rows: 1, length: 10 })]
  const pages = paginate(items, { contentHeight: 100 })
  eq('页数', pages.length, 2)
  eq('第1页节号', pages[0].sectionIndex, 0)
  eq('第2页节号', pages[1].sectionIndex, 1)
  eq('第1页页码', pages[0].pageNumber, 1)
  eq('第2页页码（重排）', pages[1].pageNumber, 1)
}

console.log('\n=== 7. 分节符：不重排时页码延续 ===')
{
  const items = [block('a', { rows: 1, length: 10 }), { t: 'break', restartNumbering: false }, block('b', { rows: 1, length: 10 })]
  const pages = paginate(items, { contentHeight: 100 })
  eq('页数', pages.length, 2)
  eq('第2页页码（延续）', pages[1].pageNumber, 2)
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
