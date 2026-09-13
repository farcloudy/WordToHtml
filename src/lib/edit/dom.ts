/**
 * 预览 DOM ⇄ 文档模型的桥。
 *
 * 预览的每一段是 contenteditable 里的一个 div（`[data-block-id]`），
 * 段落被分页切开时，同一块会在不同页上出现多片，每片带着自己的 data-from / data-to
 * （显示坐标，含自动编号前缀）。这个文件负责两件只跟浏览器有关的事：
 *
 *   1. 把一片 DOM 读回模型的 inline 序列（加粗、下划线、改色、修订、批注锚点都认）；
 *   2. 把插入符 / 选区在「显示坐标」与真实 DOM 位置之间来回换算 ——
 *      重排会重建 DOM，插入符必须靠坐标还回去，不能靠节点引用。
 *
 * 只读 DOM、不改 DOM，改 DOM 的事留给 Vue 的重新渲染。
 */

import type { Inline, RevMark } from '../types'

const TEXT_NODE = 3
const ELEMENT_NODE = 1

export interface DisplayPoint {
  blockId: string
  /** 显示坐标（含自动编号前缀） */
  offset: number
}

export interface DisplayRange {
  blockId: string
  from: number
  to: number
}

const NAMED_COLORS: Record<string, string> = {
  black: '000000',
  red: 'FF0000',
  blue: '0000FF',
  green: '008000',
  purple: '800080',
  orange: 'FF8C00',
  gray: '808080',
  grey: '808080',
  yellow: 'BF8F00',
  white: 'FFFFFF',
}

/** CSS 颜色（rgb()/hex/色名）→ 模型用的六位十六进制（不带 #） */
export function cssColorToHex(value: string): string | undefined {
  const v = value.trim().toLowerCase()
  if (v === '' || v === 'transparent' || v === 'inherit' || v === 'initial') return undefined
  const named = NAMED_COLORS[v]
  if (named) return named
  if (/^#[0-9a-f]{3}$/.test(v)) {
    return v
      .slice(1)
      .split('')
      .map((c) => `${c}${c}`)
      .join('')
      .toUpperCase()
  }
  if (/^#[0-9a-f]{6}$/.test(v)) return v.slice(1).toUpperCase()
  const m = /^rgba?\(\s*([\d.]+)[\s,]+([\d.]+)[\s,]+([\d.]+)/.exec(v)
  if (!m) return undefined
  const hex = [m[1], m[2], m[3]]
    .map((n) =>
      Math.max(0, Math.min(255, Math.round(Number(n))))
        .toString(16)
        .padStart(2, '0'),
    )
    .join('')
  return hex.toUpperCase()
}

interface InlineContext {
  bold: boolean
  underline: boolean
  color?: string
  rev?: RevMark
}

function contextOf(el: HTMLElement, inherited: InlineContext): InlineContext {
  const next: InlineContext = { ...inherited }
  if (el.tagName === 'B' || el.tagName === 'STRONG') next.bold = true
  const weight = el.style.fontWeight
  if (weight === 'bold' || weight === '700' || weight === '800' || weight === '900') {
    next.bold = true
  }
  // 只看 <u> 与行内 style：修订/批注的着色来自样式表，不该被读成下划线
  if (el.tagName === 'U') next.underline = true
  if (el.style.textDecorationLine.includes('underline')) next.underline = true
  const color = cssColorToHex(el.style.color)
  if (color) next.color = color
  if (el.tagName === 'FONT') {
    const attr = cssColorToHex(el.getAttribute('color') ?? '')
    if (attr) next.color = attr
  }
  if (el.classList.contains('wtp-rev-ins') || el.classList.contains('wtp-rev-del')) {
    next.rev = {
      kind: el.classList.contains('wtp-rev-ins') ? 'ins' : 'del',
      id: Number(el.dataset.revId ?? '0') || 0,
      author: el.dataset.revAuthor ?? '',
      date: el.dataset.revDate ?? '',
    }
  }
  return next
}

function inlineOf(text: string, ctx: InlineContext): Inline {
  return {
    t: 'text',
    text,
    ...(ctx.bold ? { bold: true } : {}),
    ...(ctx.underline ? { underline: true } : {}),
    ...(ctx.color ? { color: ctx.color } : {}),
    ...(ctx.rev ? { rev: ctx.rev } : {}),
  }
}

/** 把一片预览 DOM 读回 inline 序列。自动编号（.wtp-num）不是模型内容，跳过。 */
export function readInlines(
  root: Node,
  inherited: InlineContext = { bold: false, underline: false },
): Inline[] {
  const out: Inline[] = []
  for (const child of Array.from(root.childNodes)) {
    if (child.nodeType === TEXT_NODE) {
      const text = (child as Text).data
      if (text !== '') out.push(inlineOf(text, inherited))
      continue
    }
    if (child.nodeType !== ELEMENT_NODE) continue
    const el = child as HTMLElement
    if (el.classList.contains('wtp-num')) continue
    if (el.tagName === 'BR') continue

    if (el.classList.contains('wtp-comment')) {
      const id = Number(el.dataset.comment ?? '')
      if (Number.isFinite(id)) {
        out.push({ t: 'commentStart', commentId: id })
        out.push(...readInlines(el, contextOf(el, inherited)))
        out.push({ t: 'commentEnd', commentId: id })
        continue
      }
    }

    out.push(...readInlines(el, contextOf(el, inherited)))
  }
  return out
}

/* -------------------------------------------------------------------------- */
/* 坐标换算                                                                    */
/* -------------------------------------------------------------------------- */

/** 找到 node 所属的段落片段元素（带 data-block-id 的那一层） */
export function fragmentOf(node: Node | null): HTMLElement | null {
  let el: Node | null = node
  while (el) {
    if (el.nodeType === ELEMENT_NODE && (el as HTMLElement).dataset?.blockId) {
      return el as HTMLElement
    }
    el = el.parentNode
  }
  return null
}

/** root 内到 (target, targetOffset) 为止的字符数；兼容落点是元素边界的情况 */
export function pointToOffset(root: Node, target: Node, targetOffset: number): number {
  let count = 0
  let found = false

  const visit = (node: Node): void => {
    if (found) return
    if (node === target && node.nodeType === TEXT_NODE) {
      count += targetOffset
      found = true
      return
    }
    if (node.nodeType === TEXT_NODE) {
      count += (node as Text).data.length
      return
    }
    const children = node.childNodes
    for (let i = 0; i < children.length; i += 1) {
      if (node === target && i === targetOffset) {
        found = true
        return
      }
      visit(children[i] as Node)
      if (found) return
    }
    if (node === target) found = true
  }

  visit(root)
  return count
}

/** offsetToPoint 用到的落点类型 */
export interface CaretSpot {
  node: Node
  offset: number
}

/** 预览容器的类型（要能 querySelectorAll 又要能 contains） */
type Root = HTMLElement

/** 第 offset 个字符处的 DOM 落点 */
export function offsetToPoint(root: Node, offset: number): CaretSpot | null {
  let remaining = offset
  let last: Text | null = null

  const walk = (node: Node): CaretSpot | null => {
    if (node.nodeType === TEXT_NODE) {
      const text = node as Text
      if (remaining <= text.data.length) return { node: text, offset: remaining }
      remaining -= text.data.length
      last = text
      return null
    }
    for (const child of Array.from(node.childNodes)) {
      const hit = walk(child)
      if (hit) return hit
    }
    return null
  }

  const hit = walk(root)
  if (hit) return hit
  // 闭包里的赋值 TS 看不到，这里显式收窄一次
  const tail = last as Text | null
  if (tail) return { node: tail, offset: tail.data.length }
  return { node: root, offset: 0 }
}

/** 前缀长度（第一个片段里 .wtp-num 的字数） */
export function prefixLengthOf(frag: HTMLElement): number {
  const num = frag.querySelector('.wtp-num')
  return num ? (num.textContent ?? '').length : 0
}

function fragmentStart(frag: HTMLElement): number {
  return Number(frag.dataset.from ?? '0')
}

function fragmentEnd(frag: HTMLElement): number {
  const raw = frag.dataset.to
  if (raw !== undefined && raw !== '') return Number(raw)
  return fragmentStart(frag) + (frag.textContent ?? '').length
}

/** 光标/选区端点 → 显示坐标 */
export function displayPointOf(root: Root, node: Node, offset: number): DisplayPoint | null {
  const frag = fragmentOf(node)
  if (!frag || !root.contains(frag)) return null
  const blockId = frag.dataset.blockId ?? ''
  if (blockId === '') return null
  const inside = pointToOffset(frag, node, offset)
  const prefix = frag.querySelector('.wtp-num') ? prefixLengthOf(frag) : 0
  // 自动编号不可编辑：落点若在它之前，夹到它后面
  const local = Math.max(inside, prefix)
  return { blockId, offset: fragmentStart(frag) + local }
}

/** 找承载某个显示坐标的片段 */
export function fragmentAt(root: Root, blockId: string, offset: number): HTMLElement | null {
  const list = Array.from(
    root.querySelectorAll<HTMLElement>(`[data-block-id="${CSS.escape(blockId)}"]`),
  )
  let fallback = list[0] ?? null
  for (const frag of list) {
    if (frag.dataset.blockId !== blockId) continue
    if (offset >= fragmentStart(frag) && offset <= fragmentEnd(frag)) return frag
  }
  return fallback
}

/** 把显示坐标还原成 DOM 落点并放进选区。返回是否成功。 */
export function placeCaret(root: Root, point: DisplayPoint): boolean {
  const frag = fragmentAt(root, point.blockId, point.offset)
  if (!frag) return false
  const prefix = frag.querySelector('.wtp-num') ? prefixLengthOf(frag) : 0
  const local = Math.max(point.offset - fragmentStart(frag), prefix)
  const spot = offsetToPoint(frag, local)
  if (!spot) return false

  const host = frag.closest('[contenteditable="true"]')
  if (host instanceof HTMLElement) host.focus({ preventScroll: true })
  const sel = document.getSelection()
  if (!sel) return false
  const range = document.createRange()
  range.setStart(spot.node, spot.offset)
  range.collapse(true)
  sel.removeAllRanges()
  sel.addRange(range)
  return true
}

/** 把一段显示坐标区间还原成选区（格式化后保持选中状态用） */
export function placeRange(root: Root, from: DisplayPoint, to: DisplayPoint): boolean {
  const a = fragmentAt(root, from.blockId, from.offset)
  const b = fragmentAt(root, to.blockId, to.offset)
  if (!a || !b) return false
  const pa = offsetToPoint(a, Math.max(from.offset - fragmentStart(a), prefixLengthOf(a)))
  const pb = offsetToPoint(b, Math.max(to.offset - fragmentStart(b), prefixLengthOf(b)))
  if (!pa || !pb) return false
  const host = (a.closest('[contenteditable="true"]') ?? b) as HTMLElement
  host.focus({ preventScroll: true })
  const sel = document.getSelection()
  if (!sel) return false
  const range = document.createRange()
  range.setStart(pa.node, pa.offset)
  range.setEnd(pb.node, pb.offset)
  sel.removeAllRanges()
  sel.addRange(range)
  return true
}

/** 当前选区两端（显示坐标）。跨块选择也会如实给出两端的块。 */
export function currentRange(root: Root): { start: DisplayPoint; end: DisplayPoint } | null {
  const sel = document.getSelection()
  if (!sel || sel.rangeCount === 0) return null
  const range = sel.getRangeAt(0)
  const start = displayPointOf(root, range.startContainer, range.startOffset)
  const end = displayPointOf(root, range.endContainer, range.endOffset)
  if (!start || !end) return null
  return { start, end }
}

/**
 * 选区覆盖到的每块 [from,to)（显示坐标）。
 * 跨块、跨页都能算 —— 遍历所有片段，用 intersectsNode 判断是否落在选区内，
 * 被落点命中的那一片再按落点收紧。
 */
export function selectedRanges(root: Root): DisplayRange[] {
  const sel = document.getSelection()
  if (!sel || sel.rangeCount === 0) return []
  const range = sel.getRangeAt(0)
  if (range.collapsed) return []

  const out: DisplayRange[] = []
  for (const el of Array.from(root.querySelectorAll<HTMLElement>('[data-block-id]'))) {
    if (!range.intersectsNode(el)) continue
    const blockId = el.dataset.blockId ?? ''
    if (blockId === '') continue
    let from = fragmentStart(el)
    let to = fragmentEnd(el)
    if (el.contains(range.startContainer)) {
      from = fragmentStart(el) + pointToOffset(el, range.startContainer, range.startOffset)
    }
    if (el.contains(range.endContainer)) {
      to = fragmentStart(el) + pointToOffset(el, range.endContainer, range.endOffset)
    }
    if (to > from) out.push({ blockId, from, to })
  }
  return out
}
