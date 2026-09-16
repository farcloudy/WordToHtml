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

/**
 * 表格片段的结构性不变式（W11b）：**一片绝不重复它自己已经含有的行**。
 *
 * 续页片的重复区间是 `[0, headerTo)`、本片正文区间是 `[rowFrom, rowTo)`，两者必须不相交
 * —— 相交就是同一行被正本与重复行渲两遍（DOM 行数 ≠ 模型行数，而 docx 只写 w:tblHeader，
 * Word 的内容流里不会画两遍）。遍历所有片段，一片都不许漏。
 */
function noRepeatOverlap(label, pages) {
  const bad = []
  let seen = 0
  for (const p of pages) {
    for (const f of p.fragments) {
      if (f.rowFrom === undefined) continue
      seen += 1
      if ((f.headerTo ?? 0) > f.rowFrom) {
        bad.push(`${f.blockId} headerTo=${f.headerTo} rowFrom=${f.rowFrom}`)
      }
    }
  }
  return ok(
    `${label}：${seen} 个表格片段都满足 headerTo <= rowFrom（重复行不与本片正文重叠）`,
    seen > 0 && bad.length === 0,
    bad.join('；'),
  )
}

/** 整篇渲出的表格行去重后的条数：重复区间 [0,headerTo) 与正文区间 [rowFrom,rowTo) 都算渲出 */
function renderedRowCount(pages) {
  const rows = new Set()
  for (const p of pages) {
    for (const f of p.fragments) {
      if (f.rowFrom === undefined) continue
      for (let r = 0; r < (f.headerTo ?? 0); r += 1) rows.add(r)
      for (let r = f.rowFrom; r < f.rowTo; r += 1) rows.add(r)
    }
  }
  return rows.size
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

console.log('\n=== 3b. 片段尾标志 tail：只有覆盖到本块最后一行的那一片才带 ===')
{
  // 10 行、100 字符；版心 100px 只能放 4 行 → 三片 [0,40) / [40,80) / [80,100)
  const items = [block('long', { rows: 10, length: 100 })]
  const pages = paginate(items, { contentHeight: 100 })
  ok('第1片不到块尾（没有 tail）', pages[0].fragments[0].tail !== true)
  ok('第2片也不到块尾', pages[1].fragments[0].tail !== true)
  ok('末片带 tail', pages[2].fragments[0].tail === true)

  // 两段共用一页：tail 是「本块到没到块尾」，不是「这一页到没到末尾」
  const two = paginate(
    [block('a', { rows: 1, length: 10 }), block('b', { rows: 1, length: 10 })],
    { contentHeight: 100 },
  )
  ok('单页里的每一片都各自带 tail', two[0].fragments.every((f) => f.tail === true))
}

console.log('\n=== 3c. 尾随软换行的段落：两行、末片带 tail ===')
{
  // 「abc{br}」量出来是 2 行（第二行是空的），两行的字符起点是 0 与 3
  const item = {
    t: 'block',
    blockId: 'br',
    kind: 'body',
    displayLength: 3,
    rows: 2,
    lineHeight: 25,
    spaceBefore: 0,
    spaceAfter: 0,
    rowStarts: [0, 3],
  }
  const pages = paginate([item], { contentHeight: 100 })
  eq('页数', pages.length, 1)
  eq('一片覆盖整段', pages[0].fragments.length, 1)
  eq('片段区间仍是 0/3（软换行零宽）', `${pages[0].fragments[0].from}/${pages[0].fragments[0].to}`, '0/3')
  ok('带 tail —— 渲染时才会补占位 <br>（否则只渲 1 行、与量测差一行）', pages[0].fragments[0].tail === true)
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

console.log('\n=== 6. 分节符：新起一页 + 页码按该节设置重排 ===')
{
  const items = [
    block('a', { rows: 1, length: 10 }),
    { t: 'break', blockId: 's0', kind: 'section' },
    block('b', { rows: 1, length: 10 }),
  ]
  const pages = paginate(items, {
    contentHeight: 100,
    sections: [
      { contentHeight: 100, showPageNumber: true, restartAtOne: false },
      { contentHeight: 100, showPageNumber: true, restartAtOne: true },
    ],
  })
  eq('页数', pages.length, 2)
  eq('第1页节号', pages[0].sectionIndex, 0)
  eq('第2页节号', pages[1].sectionIndex, 1)
  eq('第1页页码', pages[0].pageNumber, 1)
  eq('第2页页码（重排）', pages[1].pageNumber, 1)
  eq('第1页末尾标记为分节符', pages[0].breaks[0]?.kind, 'section')
  eq('标记带着块 id（编辑器要按它删除）', pages[0].breaks[0]?.blockId, 's0')
  eq('第1页只有一枚标记', pages[0].breaks.length, 1)
  eq('第2页没有标记', pages[1].breaks.length, 0)
  ok('两页都显示页码', pages[0].showPageNumber === true && pages[1].showPageNumber === true)
}

console.log('\n=== 7. 分节符：不重排时页码延续 ===')
{
  const items = [
    block('a', { rows: 1, length: 10 }),
    { t: 'break', blockId: 's1', kind: 'section' },
    block('b', { rows: 1, length: 10 }),
  ]
  const pages = paginate(items, {
    contentHeight: 100,
    sections: [
      { contentHeight: 100, showPageNumber: true, restartAtOne: false },
      { contentHeight: 100, showPageNumber: true, restartAtOne: false },
    ],
  })
  eq('页数', pages.length, 2)
  eq('第2页页码（延续）', pages[1].pageNumber, 2)
}

console.log('\n=== 7.0 不传 sections：缺项按全默认兜底（显示页码、不重排）===')
{
  const items = [
    block('a', { rows: 1, length: 10 }),
    { t: 'break', blockId: 's1', kind: 'section' },
    block('b', { rows: 1, length: 10 }),
  ]
  const pages = paginate(items, { contentHeight: 100 })
  eq('页数', pages.length, 2)
  eq('第2页页码（默认不重排）', pages[1].pageNumber, 2)
  ok('默认显示页码', pages[0].showPageNumber === true && pages[1].showPageNumber === true)

  // sections 偏短（只手写了一项）时，后面的节也用同一套兜底
  const short = paginate(items, {
    contentHeight: 100,
    sections: [{ contentHeight: 100, showPageNumber: true, restartAtOne: false }],
  })
  eq('sections 偏短仍能分页', short.length, 2)
  eq('偏短时第2页页码延续', short[1].pageNumber, 2)
}

console.log('\n=== 7.0b 逐节版心高：横排节的版心更矮 → 页数变多 ===')
{
  // 8 行、行高 25px：版心 100px 每页 4 行 → 2 页；第 2 节版心只有 50px → 每页 2 行 → 4 页
  const items = [
    block('a', { rows: 4, length: 40, lineHeight: 25 }),
    { t: 'break', blockId: 's1', kind: 'section' },
    block('b', { rows: 8, length: 80, lineHeight: 25 }),
  ]
  const wide = paginate(items, {
    contentHeight: 100,
    sections: [
      { contentHeight: 100, showPageNumber: true, restartAtOne: false },
      { contentHeight: 100, showPageNumber: true, restartAtOne: false },
    ],
  })
  const short = paginate(items, {
    contentHeight: 100,
    sections: [
      { contentHeight: 100, showPageNumber: true, restartAtOne: false },
      { contentHeight: 50, showPageNumber: true, restartAtOne: false },
    ],
  })
  eq('同一版心高时 1 + 2 = 3 页', wide.length, 3)
  eq('第 2 节版心矮一半时 1 + 4 = 5 页', short.length, 5)
  eq('第 2 节第 1 页属第 2 节', short[1].sectionIndex, 1)
  eq('第 2 节每页 2 行', short[1].fragments[0].to - short[1].fragments[0].from, 20)
}

console.log('\n=== 7.0c showPageNumber 由分页器原样搬运（含「关联前节」解析后的结果）===')
{
  const items = [
    block('a', { rows: 1, length: 10 }),
    { t: 'break', blockId: 's1', kind: 'section' },
    block('b', { rows: 1, length: 10 }),
  ]
  const pages = paginate(items, {
    contentHeight: 100,
    sections: [
      { contentHeight: 100, showPageNumber: true, restartAtOne: false },
      // 第 2 节「关联前节」，继承来的结果是「不显示」
      { contentHeight: 100, showPageNumber: false, restartAtOne: false },
    ],
  })
  ok('第1页显示页码', pages[0].showPageNumber === true)
  ok('第2页不显示页码', pages[1].showPageNumber === false)
  eq('不显示页码不影响页码推进', pages[1].pageNumber, 2)
}

console.log('\n=== 7.1 分页符：只换页，不新开一节、页码连续 ===')
{
  const items = [
    block('a', { rows: 1, length: 10 }),
    { t: 'break', blockId: 'p0', kind: 'page' },
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
    { t: 'break', blockId: 'p1', kind: 'page' },
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
    { t: 'break', blockId: 'pb', kind: 'page' },
    { t: 'break', blockId: 'sc', kind: 'section' },
    block('b', { rows: 1, length: 10 }),
  ]
  const pages = paginate(items, {
    contentHeight: 100,
    sections: [
      { contentHeight: 100, showPageNumber: true, restartAtOne: false },
      { contentHeight: 100, showPageNumber: true, restartAtOne: true },
    ],
  })
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
    { t: 'break', blockId: 'sc', kind: 'section' },
    { t: 'break', blockId: 'pb', kind: 'page' },
    block('b', { rows: 1, length: 10 }),
  ]
  const pages = paginate(items, {
    contentHeight: 100,
    sections: [
      { contentHeight: 100, showPageNumber: true, restartAtOne: false },
      { contentHeight: 100, showPageNumber: true, restartAtOne: true },
    ],
  })
  eq('页数', pages.length, 2)
  eq('两枚标记都在第1页', pages[0].breaks.length, 2)
  eq('顺序保持插入顺序', pages[0].breaks.map((b) => b.kind).join(','), 'section,page')
}

console.log('\n=== 7.5 文末分节符：Word 会留一张空白页 ===')
{
  // Word 实测的 pg3：内容 + 文末分节符 → 2 页 2 节（第 2 页是空白页）
  const items = [
    block('a', { rows: 1, length: 10 }),
    { t: 'break', blockId: 'sc', kind: 'section' },
  ]
  const pages = paginate(items, {
    contentHeight: 100,
    sections: [
      { contentHeight: 100, showPageNumber: true, restartAtOne: false },
      { contentHeight: 100, showPageNumber: true, restartAtOne: true },
    ],
  })
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
    { t: 'break', blockId: 'pb', kind: 'page' },
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

console.log('\n=== 10. 表格：按行装箱（行是原子的）===')
{
  /** 造一行表格量测值（行高单位 px） */
  const row = (id, r, height) => ({ t: 'tableRow', blockId: id, row: r, height })

  // 版心 100px / 每行 25px = 每页 4 行；同页相邻同表行必须合并成一个片段（一页一张 <table>）
  const five = [0, 1, 2, 3, 4].map((r) => row('tb1', r, 25))
  const pages = paginate(five, { contentHeight: 100 })
  eq('页数', pages.length, 2)
  eq('第1页只出一个片段（同表行合并）', pages[0].fragments.length, 1)
  eq('第1页覆盖 0..4 行', `${pages[0].fragments[0].rowFrom}/${pages[0].fragments[0].rowTo}`, '0/4')
  eq('第1页不是续排', pages[0].fragments[0].continuation, false)
  eq('第2页覆盖 4..5 行', `${pages[1].fragments[0].rowFrom}/${pages[1].fragments[0].rowTo}`, '4/5')
  eq('第2页是续排（跨页必须断开）', pages[1].fragments[0].continuation, true)
  eq('表格片段没有 kind（表格不是 BlockKind）', pages[0].fragments[0].kind, undefined)
  eq('整张表共 5 行，一行都没被拆开', pages.reduce((n, p) => n + (p.fragments[0].rowTo - p.fragments[0].rowFrom), 0), 5)
}

console.log('\n=== 11. 表格行放不下就整行挪走，不吃孤行控制的副作用 ===')
{
  const row = (id, r, height) => ({ t: 'tableRow', blockId: id, row: r, height })
  // 3 行段落占 75px，剩 25px 放不下 30px 的一行 → 整行挪到第 2 页（而不是切成半行）
  const pages = paginate([block('p', { rows: 3, length: 30, lineHeight: 25 }), row('tb1', 0, 30)], {
    contentHeight: 100,
  })
  eq('页数', pages.length, 2)
  eq('第1页只有段落', pages[0].fragments.length, 1)
  eq('表格整行挪到第 2 页', pages[1].fragments[0].blockId, 'tb1')
  eq('第2页那一行是完整的', pages[1].fragments[0].rowTo - pages[1].fragments[0].rowFrom, 1)

  // 刚好放得下就放：1 行段落 + 75px 的表行 = 100px
  const fits = paginate([block('p', { rows: 1, length: 10, lineHeight: 25 }), row('tb1', 0, 75)], {
    contentHeight: 100,
  })
  eq('放得下就与段落同页', fits.length, 1)
  eq('同页两个片段（段落 + 表格）', fits[0].fragments.length, 2)

  // 差 1px 放不下：仍然整行挪走，不依赖 widow/orphan 分支的副作用
  const notFit = paginate([block('p', { rows: 1, length: 10, lineHeight: 25 }), row('tb1', 0, 76)], {
    contentHeight: 100,
  })
  eq('放不下就换页', notFit.length, 2)
  eq('第2页整行承接', `${notFit[1].fragments[0].rowFrom}/${notFit[1].fragments[0].rowTo}`, '0/1')
}

console.log('\n=== 12. 表格：不同表不合并；一页一张 <table> ===')
{
  const row = (id, r, h) => ({ t: 'tableRow', blockId: id, row: r, height: h })
  const pages = paginate(
    [row('tbA', 0, 25), row('tbA', 1, 25), row('tbB', 0, 25), row('tbB', 1, 25)],
    { contentHeight: 100 },
  )
  eq('页数', pages.length, 1)
  eq('两张表各出一个片段', pages[0].fragments.length, 2)
  eq('第一片是 tbA 的 0..2', `${pages[0].fragments[0].blockId}:${pages[0].fragments[0].rowFrom}/${pages[0].fragments[0].rowTo}`, 'tbA:0/2')
  eq('第二片是 tbB 的 0..2', `${pages[0].fragments[1].blockId}:${pages[0].fragments[1].rowFrom}/${pages[0].fragments[1].rowTo}`, 'tbB:0/2')
}

console.log('\n=== 13. 表格与段落混排、换页标记、超版心兜底 ===')
{
  const row = (id, r, h) => ({ t: 'tableRow', blockId: id, row: r, height: h })

  // 表格后面紧接着一个段落：同页继续排（表格没有段后距）
  const mixed = paginate([row('tb1', 0, 25), row('tb1', 1, 25), block('p', { rows: 1, length: 10, lineHeight: 25 })], {
    contentHeight: 100,
  })
  eq('混排同页', mixed.length, 1)
  eq('表格合并后 + 段落 = 2 片', mixed[0].fragments.length, 2)
  eq('段落那一片接着排', mixed[0].fragments[1].blockId, 'p')

  // 分页符紧跟在表格之后：标记落在表格所在页底部
  const withBreak = paginate(
    [row('tb1', 0, 25), { t: 'break', blockId: 'pgx', kind: 'page' }, block('p', { rows: 1, length: 10 })],
    { contentHeight: 100 },
  )
  eq('换页标记生效', withBreak.length, 2)
  eq('标记挂在第1页', withBreak[0].breaks[0]?.blockId, 'pgx')
  eq('第2页是新段落', withBreak[1].fragments[0].blockId, 'p')

  // 一行比整页还高：兜底放一行，不死循环
  const huge = paginate([row('tb1', 0, 500), row('tb1', 1, 25)], { contentHeight: 100 })
  eq('超版心行不死循环', huge.length, 2)
  eq('第1页放下那一行', `${huge[0].fragments[0].rowFrom}/${huge[0].fragments[0].rowTo}`, '0/1')
  eq('第2页是下一行', `${huge[1].fragments[0].rowFrom}/${huge[1].fragments[0].rowTo}`, '1/2')
}

console.log('\n=== 14. 重复标题行（w:tblHeader）：续页片带 headerTo、并按重复行高扣可用高度 ===')
{
  const row = (id, r, height, extra = {}) => ({ t: 'tableRow', blockId: id, row: r, height, ...extra })
  /** 6 行 / 每行 25px / 版心 100px = 每页 4 行；前 1 行是标题行（25px） */
  const repeat1 = { repeatRows: 1, headerHeight: 25 }

  // ---- 不标 headerRows：与改动前逐字节相同（片段上不该多出 headerTo） ----
  const plainPages = paginate([0, 1, 2, 3, 4, 5].map((r) => row('tb1', r, 25)), { contentHeight: 100 })
  eq('不标：页数', plainPages.length, 2)
  eq('不标：第1页 0/4', `${plainPages[0].fragments[0].rowFrom}/${plainPages[0].fragments[0].rowTo}`, '0/4')
  eq('不标：第2页 4/6', `${plainPages[1].fragments[0].rowFrom}/${plainPages[1].fragments[0].rowTo}`, '4/6')
  ok('不标：第一片不带 headerTo', plainPages[0].fragments[0].headerTo === undefined)
  ok('不标：续页片也不带 headerTo（默认值不落字段）', plainPages[1].fragments[0].headerTo === undefined)

  // ---- 标 1 行：第一片不带 headerTo；续页让出 25px → 每页只剩 3 行 ----
  const markedPages = paginate(
    [0, 1, 2, 3, 4, 5, 6, 7, 8, 9].map((r) => row('tb1', r, 25, repeat1)),
    { contentHeight: 100 },
  )
  const noRepeat = paginate([0, 1, 2, 3, 4, 5, 6, 7, 8, 9].map((r) => row('tb1', r, 25)), {
    contentHeight: 100,
  })
  eq('标 1 行：页数（4 + 3 + 3）', markedPages.length, 3)
  eq(
    '同样 10 行不标时是 4 + 4 + 2（标了以后每页少一行，正是重复行占掉的高度）',
    noRepeat.map((p) => p.fragments[0].rowTo).join(','),
    '4,8,10',
  )
  ok('标 1 行：第一片不带 headerTo（它本来就是第一页，重复行还没有意义）', markedPages[0].fragments[0].headerTo === undefined)
  eq('标 1 行：第1页仍放 4 行', `${markedPages[0].fragments[0].rowFrom}/${markedPages[0].fragments[0].rowTo}`, '0/4')
  eq('标 1 行：第2页的区间', `${markedPages[1].fragments[0].rowFrom}/${markedPages[1].fragments[0].rowTo}`, '4/7')
  eq('标 1 行：第1页续页片带 headerTo=1', markedPages[1].fragments[0].headerTo, 1)
  eq('标 1 行：第3页的区间', `${markedPages[2].fragments[0].rowFrom}/${markedPages[2].fragments[0].rowTo}`, '7/10')
  eq('标 1 行：第2页续页片也带 headerTo=1', markedPages[2].fragments[0].headerTo, 1)
  ok(
    '续页片「正文行高 + headerHeight」正好装满版心（可用高度真的少了那一行）',
    markedPages.slice(1).every((p) => {
      const f = p.fragments[0]
      const h = (f.rowTo - f.rowFrom) * 25 + 25
      return h === 100
    }),
  )
  eq(
    '整张表 10 行一行不多一行不少',
    markedPages.reduce((n, p) => n + (p.fragments[0].rowTo - p.fragments[0].rowFrom), 0),
    10,
  )

  // ---- 表的第一片就是新页（上一块把它挤下去）：仍不带 headerTo ----
  const pushed = paginate(
    [block('p', { rows: 1, length: 10, lineHeight: 76 }), ...[0, 1].map((r) => row('tb1', r, 25, repeat1))],
    { contentHeight: 100 },
  )
  eq('表被挤到第 2 页', pushed.length, 2)
  ok(
    '表的第一片不带 headerTo —— 它在页顶也是从第 0 行开始，没有「重复」可言',
    pushed[1].fragments[0].headerTo === undefined,
  )

  // ---- 「重复行 + 一行正文」都放不下：按兜底硬放，不死循环 ----
  const tiny = paginate([0, 1, 2].map((r) => row('tb1', r, 25, repeat1)), { contentHeight: 30 })
  eq('版心只剩 30px 也不死循环（一行一页）', tiny.length, 3)
  eq('第1页放第 0 行', `${tiny[0].fragments[0].rowFrom}/${tiny[0].fragments[0].rowTo}`, '0/1')
  ok('第2页仍是续页（带 headerTo）', tiny[1].fragments[0].headerTo === 1)
  eq('第3页放最后一行', `${tiny[2].fragments[0].rowFrom}/${tiny[2].fragments[0].rowTo}`, '2/3')

  // ---- 遍历上面每一组片段：结构性不变式「一片绝不重复它自己含有的行」 ----
  noRepeatOverlap('不标', plainPages)
  noRepeatOverlap('标 1 行', markedPages)
  noRepeatOverlap('表被挤到新页', pushed)
  noRepeatOverlap('版心只剩 30px', tiny)

  // ---- W11b-1：起表处只剩 1 行空间、headerRows=2 → 整张表挪到下一页（首片含足 2 行）----
  // 版心 100：段落占 60 → 剩 40，只放得下 1 行；而标题块 = 前 2 行 = 50px 放不下。
  // 旧算法就这么起表 → 首片只有第 0 行 → 续页片带 headerTo=2 把第 1 行渲两遍（验收 ⑩）。
  const repeat2 = { repeatRows: 2, headerHeight: 50 }
  const keptHeader = paginate(
    [
      block('p', { rows: 1, length: 10, lineHeight: 60 }),
      ...[0, 1, 2, 3].map((r) => row('tb1', r, 25, repeat2)),
    ],
    { contentHeight: 100 },
  )
  eq('起表：整张表挪到下一页（第1页只剩那段文字）', keptHeader.length, 2)
  eq('起表：第1页只有 1 片，是段落不是表格', keptHeader[0].fragments.length, 1)
  ok('起表：第1页那片是段落片段', keptHeader[0].fragments[0].rowFrom === undefined)
  eq(
    '起表：第2页首片从第 0 行起、4 行整整齐齐都在这一页',
    `${keptHeader[1].fragments[0].rowFrom}/${keptHeader[1].fragments[0].rowTo}`,
    '0/4',
  )
  ok(
    '起表：首片含足整个标题块（>= headerRows = 2 行）',
    keptHeader[1].fragments[0].rowTo - keptHeader[1].fragments[0].rowFrom >= 2,
  )
  ok('起表：首片是「第一片」，不带重复行', keptHeader[1].fragments[0].headerTo === undefined)
  ok('起表：第1页不再有表格碎片（没有「只有标题行第 0 行」那种首片）', keptHeader[0].fragments.every((f) => f.rowFrom === undefined))
  noRepeatOverlap('起表保住标题块', keptHeader)

  // ---- W11b-2 退化情形：headerRows=2 但标题块高过整页 → 起表那次换页救不了，
  //      「不许重叠」是唯一的安全网（旧算法这里会把第 1 行渲两遍）----
  // 版心 100、行高 80/80/20/20 → 标题块 = 160 > 100
  const tallHeader = { repeatRows: 2, headerHeight: 160 }
  const degenerate = paginate([0, 1, 2, 3].map((r) => row('tb1', r, r < 2 ? 80 : 20, tallHeader)), {
    contentHeight: 100,
  })
  noRepeatOverlap('退化（标题块高过整页）', degenerate)
  eq(
    '退化：整篇渲出的行去重后 = 模型 4 行（不多不少、没有幻影行）',
    renderedRowCount(degenerate),
    4,
  )
  eq(
    '退化：正文区间合起来正好覆盖模型每一行（一行都不丢）',
    degenerate.reduce(
      (n, p) => n + p.fragments.filter((f) => f.rowFrom !== undefined).reduce((m, f) => m + (f.rowTo - f.rowFrom), 0),
      0,
    ),
    4,
  )
  eq(
    '退化：续页片的重复行也不越过「本片正文起点」（第 1 行只重复第 0 行）',
    degenerate[1].fragments[0].headerTo,
    1,
  )
}

console.log(`\n通过 ${passed} 项，失败 ${failed} 项`)
process.exit(failed === 0 ? 0 : 1)
