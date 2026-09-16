/**
 * DOM 测量：把每个段落渲染进隐藏容器，量出真实行数、行高、以及每一行开头
 * 对应的字符偏移。分页算法拿到的就是这些实测值。
 *
 * 之所以用「行盒实测」而不是用「字符宽度累加算」：两端对齐、中西文混排、
 * 标点避头尾这些都会影响换行位置，只有让浏览器真的排一遍才是准的。
 */

import { contentBoxPx, lineSpacePt, ptToPx } from '../spec'
import type { Spec } from '../spec'
import { computeNumbering } from '../numbering'
import { resolveSections } from '../section'
import type { DocModel, TableBlock } from '../types'
import { headerRowCount } from '../types'
import { renderInlinesHtml, renderTableFragment } from './html'
import type { MeasuredBlock, MeasuredItem, MeasuredTableRow } from './paginate'

interface TextPos {
  node: Text
  start: number
  end: number
}

function collectTextPositions(el: HTMLElement): TextPos[] {
  const out: TextPos[] = []
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT)
  let cursor = 0
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const text = node as Text
    const len = text.data.length
    if (len === 0) continue
    out.push({ node: text, start: cursor, end: cursor + len })
    cursor += len
  }
  return out
}

function positionAt(
  positions: readonly TextPos[],
  offset: number,
): { node: Text; offset: number } | null {
  for (const p of positions) {
    if (offset >= p.start && offset <= p.end) {
      return { node: p.node, offset: offset - p.start }
    }
  }
  return null
}

/** 某个字符所在行盒的顶端坐标（px） */
function topOfChar(range: Range, positions: readonly TextPos[], offset: number): number {
  const a = positionAt(positions, offset)
  if (!a) return Number.NaN
  const b = positionAt(positions, offset + 1)
  range.setStart(a.node, a.offset)
  range.setEnd(b ? b.node : a.node, b ? b.offset : a.offset)
  if (range.collapsed) {
    range.setEnd(a.node, Math.min(a.offset + 1, a.node.data.length))
  }
  const rects = range.getClientRects()
  return rects.length > 0 ? (rects[0]?.top ?? Number.NaN) : Number.NaN
}

/**
 * 用二分查找定位每一行的起始字符。
 * 逐字符线性扫描太慢（长段落几千个字符会卡住），二分是 O(行数 × log 字数)。
 */
function findRowStarts(el: HTMLElement, rowTops: readonly number[], total: number): number[] {
  const positions = collectTextPositions(el)
  if (positions.length === 0 || total === 0) return [0]

  const range = document.createRange()
  const starts = [0]
  let lo = 0

  for (let row = 1; row < rowTops.length; row += 1) {
    const target = rowTops[row] ?? 0
    let low = lo
    let high = total
    // 找出第一个 top 不小于该行顶端的位置
    while (low < high) {
      const mid = (low + high) >> 1
      const top = topOfChar(range, positions, mid)
      if (Number.isNaN(top) || top < target - 0.5) low = mid + 1
      else high = mid
    }
    starts.push(low)
    lo = low
  }

  return starts
}

/**
 * 把矩形按 top 归并成「行」。
 *
 * 关键：Range.getClientRects() 返回的是**每个文本节点片段一个矩形**，不是每行一个。
 * 段落里只要有 <b>、改色、修订、批注这些内联元素，同一行就会被切成多个矩形。
 * 不归并就会高估行数 —— 实测曾把 3 行的段落量成 10 行，导致整页算胖、
 * 提前断页并在页底留下大片空白。行距恒大于 1px，所以 1px 容差足够区分相邻行。
 */
function groupLineTops(rects: readonly DOMRect[]): number[] {
  const tops: number[] = []
  for (const rect of rects) {
    const last = tops[tops.length - 1]
    if (last === undefined || rect.top - last > 1) tops.push(rect.top)
  }
  return tops
}

function measureElement(
  el: HTMLElement,
  spec: Spec,
  kind: MeasuredBlock['kind'],
  blockId: string,
  displayLength: number,
): MeasuredBlock {
  const s = spec.styles[kind]
  const height = el.getBoundingClientRect().height

  const range = document.createRange()
  range.selectNodeContents(el)
  const rects = Array.from(range.getClientRects()).filter((r) => r.height > 0)
  const lineTops = groupLineTops(rects)
  const rows = Math.max(1, lineTops.length)

  // 多行时用首末行顶端的平均间距，比取相邻两行更稳（atLeast 行距下各行可能微异）
  const lineHeight =
    rows > 1 ? ((lineTops[rows - 1] ?? 0) - (lineTops[0] ?? 0)) / (rows - 1) : height

  const rowStarts = rows > 1 ? findRowStarts(el, lineTops, displayLength) : [0]

  return {
    t: 'block',
    blockId,
    kind,
    displayLength,
    rows,
    lineHeight,
    spaceBefore: ptToPx(lineSpacePt(s.spaceBeforeLines, spec)),
    spaceAfter: ptToPx(lineSpacePt(s.spaceAfterLines, spec)),
    rowStarts,
  }
}

/**
 * 量测缓存：块 id → 上次的签名与结果。
 *
 * 编辑时每次敲键都要重排判断「分页有没有变」，但整篇重新量测太贵
 * （长段落要跑二分找行首，每次 getClientRects 都会触发布局）。
 * 块的排版只取决于「版心宽度 + 本块内容 + 本块样式」，
 * 所以内容没变的块可以直接复用上次的结果，只量改过的那一块。
 * 签名里带上渲出来的 HTML，等于把「文字、加粗、颜色、修订、编号」一起算了进去
 * —— 表格的签名含**整张表**的 HTML，任一格改动即失效。
 *
 * **签名里还带版心宽度**（`<kind>\0<width>\0<html>`）：W5 起每节可以有各自的页面方向，
 * 版心宽因此不再全局固定。把宽度写进签名，「换方向 / 换模板就必须清缓存」这条
 * 不再是靠人记得的约定，而是签名自己保证的（宽度变了签名必然不等）。
 * clearMeasureCache 仍保留（处理「规格表整体换了一套」这类改动）。
 */
export type MeasuredCacheValue = MeasuredBlock | MeasuredTableRow[]

export interface MeasureCacheEntry {
  signature: string
  measured: MeasuredCacheValue
}

export type MeasureCache = Map<string, MeasureCacheEntry>

/** 清空缓存（整份规格表换了一套时调用；单块失效靠签名里带的宽度与 HTML 自动完成） */
export function clearMeasureCache(cache: MeasureCache): void {
  cache.clear()
}

/**
 * 测量整篇文档（带增量缓存）。
 *
 * root 只需是一个已挂载在文档里的元素（用来挂载测量容器）。
 * 探针宽度**逐节切换**：横排节的版心宽与纵排节不同，量出来的行盒也就不同。
 * 切换时必须先把上一节的矩形读完（改宽会让已挂进去的元素重排，rect 就变了），
 * 所以结构是「遇到分节符 → flush 上一批 → 再改宽度」。
 */
export function measureDocument(
  doc: DocModel,
  spec: Spec,
  root: HTMLElement,
  cache?: MeasureCache,
): MeasuredItem[] {
  const numbering = computeNumbering(doc.blocks, (b) =>
    b.t === 'textBlock' ? spec.styles[b.kind].numbering : 'none',
  )
  const sections = resolveSections(doc, spec)
  const widthOf = (index: number): number =>
    (sections[index] ?? sections[0])?.content.width ?? contentBoxPx(spec).width

  const probe = document.createElement('div')
  probe.className = 'wtp-probe'

  const pending: {
    el: HTMLElement
    kind: MeasuredBlock['kind']
    id: string
    length: number
    signature: string
  }[] = []

  const pendingTables: {
    el: HTMLElement
    block: TableBlock
    signature: string
  }[] = []

  const resolved = new Map<string, MeasuredBlock>()
  const resolvedTables = new Map<string, MeasuredTableRow[]>()
  const alive = new Set<string>()

  const flush = (): void => {
    if (pending.length === 0 && pendingTables.length === 0) return
    root.appendChild(probe)
    // 强制一次布局，后面读 rect 就不会反复触发回流
    void probe.getBoundingClientRect()

    for (const item of pending) {
      const measured = measureElement(item.el, spec, item.kind, item.id, item.length)
      resolved.set(item.id, measured)
      cache?.set(item.id, { signature: item.signature, measured })
    }
    for (const item of pendingTables) {
      const rows = measureTableRows(item.el, item.block)
      resolvedTables.set(item.block.id, rows)
      cache?.set(item.block.id, { signature: item.signature, measured: rows })
    }

    root.removeChild(probe)
    // 探针会被下一节复用，上一节的元素必须清掉 —— 留着会让下一节的探针里
    // 混进上一节的内容，量出来的行盒全错（而且宽度是新的，重排结果无从预料）
    probe.replaceChildren()
    pending.length = 0
    pendingTables.length = 0
  }

  let sectionIndex = 0
  probe.style.width = `${widthOf(0)}px`

  for (const block of doc.blocks) {
    if (block.t === 'sectionBreak') {
      // 先把上一节的矩形读完，再改探针宽度 —— 反了的话量到的是改宽后重排过的版本
      flush()
      sectionIndex += 1
      probe.style.width = `${widthOf(sectionIndex)}px`
      continue
    }

    const width = widthOf(sectionIndex)

    if (block.t === 'table') {
      // 表格走独立分支：整张表渲进探针，只量每行的实测高（行是原子的，不量 rowStarts）
      alive.add(block.id)
      const html = renderTableFragment(block, 0, block.rows.length)
      // 签名带上 headerRows：改「重复标题行」一个字都不改渲出来的 HTML（重复行不在量测路径里），
      // 漏进签名就会缓存命中 → 分页仍按旧的 headerHeight 记账（radio 点了没反应）
      const signature = `table\u0000${width}\u0000${headerRowCount(block)}\u0000${html}`
      const cached = cache?.get(block.id)
      if (cached && cached.signature === signature && Array.isArray(cached.measured)) {
        resolvedTables.set(block.id, cached.measured)
        continue
      }
      const el = document.createElement('div')
      el.className = 'wtp-tableFrag'
      el.innerHTML = html
      probe.appendChild(el)
      pendingTables.push({ el, block, signature })
      continue
    }
    if (block.t !== 'textBlock') continue
    alive.add(block.id)
    const prefix = numbering.get(block.id) ?? ''
    const html = renderInlinesHtml(block.inlines, prefix)
    const signature = `${block.kind}\u0000${width}\u0000${html}`

    let length = prefix.length
    for (const inline of block.inlines) {
      if (inline.t === 'text') length += inline.text.length
    }

    const cached = cache?.get(block.id)
    if (cached && cached.signature === signature && !Array.isArray(cached.measured)) {
      resolved.set(block.id, cached.measured)
      continue
    }

    const el = document.createElement('div')
    el.className = `wtp-${block.kind}`
    el.innerHTML = html
    probe.appendChild(el)
    pending.push({ el, kind: block.kind, id: block.id, length, signature })
  }

  flush()

  if (cache) {
    for (const id of cache.keys()) {
      if (!alive.has(id)) cache.delete(id)
    }
  }

  return doc.blocks.flatMap<MeasuredItem>((block) => {
    if (block.t === 'sectionBreak') {
      return [{ t: 'break', blockId: block.id, kind: 'section' }]
    }
    if (block.t === 'pageBreak') {
      return [{ t: 'break', blockId: block.id, kind: 'page' }]
    }
    if (block.t === 'table') return resolvedTables.get(block.id) ?? []
    const m = resolved.get(block.id)
    return m ? [m] : []
  })
}

/** 一张表里每一行的实测高（px）。行序就是 block.rows 的顺序 */
function measureTableRows(el: HTMLElement, block: TableBlock): MeasuredTableRow[] {
  const rows = Array.from(el.querySelectorAll('tr')).map((tr, row) => ({
    t: 'tableRow' as const,
    blockId: block.id,
    row,
    height: tr.getBoundingClientRect().height,
  }))
  /*
   * 重复标题行：前 N 行的实测高之和（0 = 不重复）。**逐行加起来现算**而不是另量一遍 ——
   * 探针里渲的是整张表（重复行不进量测路径），这 N 行的高度上面刚量到。
   * 同一张表每一行都带同一对值，分页侧不必自己攒状态。
   * 归零时两个字段都不落（与「默认值不落模型」同一条约定，也让默认表的量测值逐字节不变）。
   */
  const repeat = headerRowCount(block)
  if (repeat <= 0) return rows
  let headerHeight = 0
  for (let r = 0; r < repeat; r += 1) headerHeight += rows[r]?.height ?? 0
  return rows.map((row) => ({ ...row, repeatRows: repeat, headerHeight }))
}
