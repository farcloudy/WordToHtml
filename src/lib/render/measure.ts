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
import type { DocModel } from '../types'
import { renderInlinesHtml } from './html'
import type { MeasuredBlock, MeasuredItem } from './paginate'

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
 * 块的排版只取决于「排版宽度（全局固定）+ 本块内容 + 本块样式」，
 * 所以内容没变的块可以直接复用上次的结果，只量改过的那一块。
 * 签名里带上渲出来的 HTML，等于把「文字、加粗、颜色、修订、编号」一起算了进去。
 */
export interface MeasureCacheEntry {
  signature: string
  measured: MeasuredBlock
}

export type MeasureCache = Map<string, MeasureCacheEntry>

/** 清空缓存（规格表换了一套样式时必须调用，否则会拿旧样式的量测值） */
export function clearMeasureCache(cache: MeasureCache): void {
  cache.clear()
}

/**
 * 测量整篇文档（带增量缓存）。
 *
 * root 只需是一个已挂载在文档里的元素（用来挂载测量容器）；
 * 测量容器的宽度由版心尺寸决定，与预览页的版心一致。
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

  const probe = document.createElement('div')
  probe.className = 'wtp-probe'
  probe.style.width = `${contentBoxPx(spec).width}px`

  const pending: {
    el: HTMLElement
    kind: MeasuredBlock['kind']
    id: string
    length: number
    signature: string
  }[] = []

  const resolved = new Map<string, MeasuredBlock>()
  const alive = new Set<string>()

  for (const block of doc.blocks) {
    if (block.t !== 'textBlock') continue
    alive.add(block.id)
    const prefix = numbering.get(block.id) ?? ''
    const html = renderInlinesHtml(block.inlines, prefix)
    const signature = `${block.kind}\u0000${html}`

    let length = prefix.length
    for (const inline of block.inlines) {
      if (inline.t === 'text') length += inline.text.length
    }

    const cached = cache?.get(block.id)
    if (cached && cached.signature === signature) {
      resolved.set(block.id, cached.measured)
      continue
    }

    const el = document.createElement('div')
    el.className = `wtp-${block.kind}`
    el.innerHTML = html
    probe.appendChild(el)
    pending.push({ el, kind: block.kind, id: block.id, length, signature })
  }

  if (cache) {
    for (const id of cache.keys()) {
      if (!alive.has(id)) cache.delete(id)
    }
  }

  if (pending.length > 0) {
    root.appendChild(probe)
    // 强制一次布局，后面读 rect 就不会反复触发回流
    void probe.getBoundingClientRect()

    for (const item of pending) {
      const measured = measureElement(item.el, spec, item.kind, item.id, item.length)
      resolved.set(item.id, measured)
      cache?.set(item.id, { signature: item.signature, measured })
    }

    root.removeChild(probe)
  }

  return doc.blocks.flatMap<MeasuredItem>((block) => {
    if (block.t === 'sectionBreak') {
      return [
        {
          t: 'break',
          blockId: block.id,
          kind: 'section',
          restartNumbering: block.restartNumbering,
        },
      ]
    }
    if (block.t === 'pageBreak') {
      return [{ t: 'break', blockId: block.id, kind: 'page', restartNumbering: false }]
    }
    const m = resolved.get(block.id)
    return m ? [m] : []
  })
}
