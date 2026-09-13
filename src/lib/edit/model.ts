/**
 * 编辑操作的模型层（纯函数，不碰 DOM，可在 node 里单测）。
 *
 * 编辑采用「DOM 是手感的真相、模型是导出的真相」这一分工：
 * 浏览器负责把字敲进去（原生编辑、输入法、原生选区都留着），
 * 每次 input 再把 DOM 读回模型；结构性操作（回车、退格合并、工具栏格式化、
 * 批注）反过来直接改模型，再重排渲染。两者都收敛到这一层，
 * 因此 docx 导出的永远是模型，不会出现「界面改了、导出没改」。
 *
 * 坐标系：这里的偏移一律是**模型文字坐标**（不含标题自动编号前缀）。
 * 预览 DOM 用的是「显示坐标」（含前缀），换算只在 dom.ts 里做。
 */

import type { BlockKind } from '../spec'
import type { Block, CommentDef, DocModel, Inline, RevMark, TextBlock } from '../types'
import { nextBlockId, plainText } from '../types'

export interface BlockPoint {
  blockId: string
  /** 模型文字坐标下的字符偏移 */
  offset: number
}

/** 工具栏要的选区信息（模型文字坐标，不含自动编号前缀） */
export interface EditorSelection {
  blockId: string
  kind: BlockKind
  from: number
  to: number
  /** 是否只是一个插入符（没有选中文字） */
  collapsed: boolean
  /** 选区内文字是否全加粗 */
  bold: boolean
  /** 选区内文字是否同色；混色或无色时为 undefined */
  color?: string
}

export function findBlock(doc: DocModel, id: string): TextBlock | undefined {
  for (const block of doc.blocks) {
    if (block.t === 'textBlock' && block.id === id) return block
  }
  return undefined
}

export function findBlockIndex(doc: DocModel, id: string): number {
  return doc.blocks.findIndex((block) => block.t === 'textBlock' && block.id === id)
}

export function blockLength(block: TextBlock): number {
  return plainText(block).length
}

/**
 * 按字符区间切 inline 片段，批注锚点跟着它夹住的那段文字走。
 *
 * 与 types.ts 的 sliceInlines 的差别：那个是给分页用的 —— 它无条件保留所有
 * 批注锚点，好让跨页的每个片段都能把高亮画全。编辑读回时不能那样：
 * 把 [from,to) 的片段替换掉以后，[0,from) 与 [to,len) 里残留的锚点会变成
 * 重复的批注标记。所以这里按锚点的位置过滤：起点落在 [from,to) 内才留，
 * 终点落在 (from,to] 内才留。
 */
export function sliceStrict(inlines: readonly Inline[], from: number, to: number): Inline[] {
  const out: Inline[] = []
  let cursor = 0
  for (const inline of inlines) {
    if (inline.t === 'commentStart') {
      if (cursor >= from && cursor < to) out.push(inline)
      continue
    }
    if (inline.t === 'commentEnd') {
      if (cursor > from && cursor <= to) out.push(inline)
      continue
    }
    const start = cursor
    const end = cursor + inline.text.length
    cursor = end
    if (end <= from || start >= to) continue
    const cutStart = Math.max(from, start) - start
    const cutEnd = Math.min(to, end) - start
    out.push({ ...inline, text: inline.text.slice(cutStart, cutEnd) })
  }
  return out
}

/**
 * 用新片段替换块内 [from,to) 的文字（其余部分原样保留）。
 *
 * 批注锚点不能像 sliceStrict 那样按区间一刀切：夹住被替换文字的锚点，起点必须留在
 * 替换内容之前、终点必须留在之后。否则会剩下一个没有终点的 commentStart，
 * 渲染时它会把这一段余下的文字全吞进高亮里，导出 docx 也会写出没闭合的批注范围。
 * 所以中间的锚点按「夹到边界」处理：起点并入前半、终点并入后半，新内容自然被包住。
 * 两条规则（前半 cursor < a、后半 cursor >= b）互斥，插入（from == to）时也不会重复。
 */
export function replaceRange(
  block: TextBlock,
  from: number,
  to: number,
  inlines: readonly Inline[],
): void {
  const len = blockLength(block)
  const a = Math.max(0, Math.min(from, len))
  const b = Math.max(a, Math.min(to, len))
  const out: Inline[] = []

  let cursor = 0
  for (const inline of block.inlines) {
    if (inline.t !== 'text') {
      if (cursor < a || (cursor < b && inline.t === 'commentStart')) out.push(inline)
      continue
    }
    const start = cursor
    cursor += inline.text.length
    if (start >= a) continue
    const text = inline.text.slice(0, Math.min(a, cursor) - start)
    if (text !== '') out.push({ ...inline, text })
  }

  out.push(...inlines)

  cursor = 0
  for (const inline of block.inlines) {
    if (inline.t !== 'text') {
      if (cursor >= b || (cursor > a && cursor < b && inline.t === 'commentEnd')) {
        out.push(inline)
      }
      continue
    }
    const start = cursor
    cursor += inline.text.length
    if (cursor <= b) continue
    out.push({ ...inline, text: inline.text.slice(Math.max(b, start) - start) })
  }

  block.inlines = out
}

/** 在 offset 处插入纯文本（可选地带上修订标记） */
export function insertText(
  doc: DocModel,
  blockId: string,
  offset: number,
  text: string,
  rev?: RevMark,
): void {
  const block = findBlock(doc, blockId)
  if (!block || text === '') return
  const piece: Inline = { t: 'text', text, ...(rev ? { rev } : {}) }
  replaceRange(block, offset, offset, [piece])
}

/**
 * 删除 [from,to)。`rev` 非空时不真删，而是把删掉的文字标成 w:del 留在原处
 * ——这正是 Word 的「修订模式」在界面上看到的效果。
 */
export function deleteRange(
  doc: DocModel,
  blockId: string,
  from: number,
  to: number,
  rev?: RevMark,
): void {
  const block = findBlock(doc, blockId)
  if (!block || to <= from) return
  if (!rev) {
    replaceRange(block, from, to, [])
    return
  }
  // 保留原文字，只把区间内的 text inline 盖上删除标记（原来的 ins 标记要让位）
  const kept = sliceStrict(block.inlines, from, to).map((inline): Inline => {
    if (inline.t !== 'text') return inline
    return { ...inline, rev: { ...rev } }
  })
  replaceRange(block, from, to, kept)
}

/** 把块在 offset 处切成两块，后半段成为 tailKind 的新块；返回新块 id */
export function splitBlock(
  doc: DocModel,
  blockId: string,
  offset: number,
  tailKind: BlockKind,
): string {
  const index = findBlockIndex(doc, blockId)
  const block = doc.blocks[index]
  if (index < 0 || !block || block.t !== 'textBlock') return ''
  const len = blockLength(block)
  const at = Math.max(0, Math.min(offset, len))
  const tail = sliceStrict(block.inlines, at, len)
  block.inlines = sliceStrict(block.inlines, 0, at)
  const created: TextBlock = { t: 'textBlock', id: nextBlockId(), kind: tailKind, inlines: tail }
  doc.blocks.splice(index + 1, 0, created)
  return created.id
}

/** 把当前块并入前一块；返回合并后的落点（前一块的末尾）。前一块是分节符时不合并。 */
export function mergeIntoPrevious(doc: DocModel, blockId: string): BlockPoint | null {
  const index = findBlockIndex(doc, blockId)
  const block = doc.blocks[index]
  const prev = doc.blocks[index - 1]
  if (index <= 0 || !block || !prev || block.t !== 'textBlock' || prev.t !== 'textBlock') {
    return null
  }
  const at = blockLength(prev)
  prev.inlines = [...prev.inlines, ...block.inlines]
  doc.blocks.splice(index, 1)
  return { blockId: prev.id, offset: at }
}

/** 删除一个块（退格删空段、或块被并入后清理用） */
export function removeBlock(doc: DocModel, blockId: string): void {
  const index = findBlockIndex(doc, blockId)
  if (index >= 0) doc.blocks.splice(index, 1)
}

export function setBlockKind(doc: DocModel, blockId: string, kind: BlockKind): void {
  const block = findBlock(doc, blockId)
  if (block) block.kind = kind
}

/** 判断 [from,to) 内的文字是否全部加粗（空区间返回 false） */
export function rangeIsBold(block: TextBlock, from: number, to: number): boolean {
  if (to <= from) return false
  let seen = false
  let cursor = 0
  for (const inline of block.inlines) {
    if (inline.t !== 'text') continue
    const start = cursor
    const end = cursor + inline.text.length
    cursor = end
    if (end <= from || start >= to) continue
    seen = true
    if (!inline.bold) return false
  }
  return seen
}

/** 判断 [from,to) 内的文字是否全是同一个颜色；不一致返回 undefined */
export function rangeColor(block: TextBlock, from: number, to: number): string | undefined {
  if (to <= from) return undefined
  let found: string | undefined
  let seen = false
  let cursor = 0
  for (const inline of block.inlines) {
    if (inline.t !== 'text') continue
    const start = cursor
    const end = cursor + inline.text.length
    cursor = end
    if (end <= from || start >= to) continue
    if (!seen) {
      seen = true
      found = inline.color
    } else if (found !== inline.color) {
      return undefined
    }
  }
  return seen ? found : undefined
}

/**
 * 给 [from,to) 套格式。颜色传 null 表示「恢复默认色」（清掉 color）。
 * 加粗传 true/false，未传的字段保持原样。
 */
export function applyFormat(
  doc: DocModel,
  blockId: string,
  from: number,
  to: number,
  patch: { bold?: boolean; color?: string | null },
): void {
  const block = findBlock(doc, blockId)
  if (!block || to <= from) return
  replaceRange(
    block,
    from,
    to,
    sliceStrict(block.inlines, from, to).map((inline): Inline => {
      if (inline.t !== 'text') return inline
      const next: Inline = { ...inline }
      if (patch.bold !== undefined) {
        if (patch.bold) next.bold = true
        else delete next.bold
      }
      if (patch.color !== undefined) {
        if (patch.color) next.color = patch.color
        else delete next.color
      }
      return next
    }),
  )
}

/** 给 [from,to) 加一条批注；返回新批注 id */
export function addComment(
  doc: DocModel,
  blockId: string,
  from: number,
  to: number,
  text: string,
  author: string,
  date: string,
): number {
  const block = findBlock(doc, blockId)
  if (!block || to <= from) return -1
  const id = doc.comments.reduce((max, c) => Math.max(max, c.id), -1) + 1
  const def: CommentDef = { id, author, date, text }
  doc.comments.push(def)
  // 用 replaceRange 落锚点：它会顺带把与新区间相交的旧锚点配对好（见该函数的说明）
  replaceRange(block, from, to, [
    { t: 'commentStart', commentId: id },
    ...sliceStrict(block.inlines, from, to),
    { t: 'commentEnd', commentId: id },
  ])
  return id
}

/** 回复某条批注（Word 里表现为同一线程下的追加） */
export function replyComment(
  doc: DocModel,
  parentId: number,
  text: string,
  author: string,
  date: string,
): number {
  const id = doc.comments.reduce((max, c) => Math.max(max, c.id), -1) + 1
  doc.comments.push({ id, author, date, text, parentId })
  return id
}

/** 删除整条批注及其锚点（父批注连同回复一起删） */
export function removeComment(doc: DocModel, commentId: number): void {
  const doomed = new Set<number>([commentId])
  let grew = true
  while (grew) {
    grew = false
    for (const c of doc.comments) {
      if (c.parentId !== undefined && doomed.has(c.parentId) && !doomed.has(c.id)) {
        doomed.add(c.id)
        grew = true
      }
    }
  }
  doc.comments = doc.comments.filter((c) => !doomed.has(c.id))
  for (const block of doc.blocks) {
    if (block.t !== 'textBlock') continue
    block.inlines = block.inlines.filter(
      (inline) =>
        (inline.t !== 'commentStart' && inline.t !== 'commentEnd') || !doomed.has(inline.commentId),
    )
  }
}

/** 深拷贝模型（撤销栈、渲染快照都用它，避免共享引用被后续编辑改到） */
export function cloneDoc(doc: DocModel): DocModel {
  return {
    blocks: doc.blocks.map((block): Block => {
      if (block.t === 'sectionBreak') return { ...block }
      return {
        t: 'textBlock',
        id: block.id,
        kind: block.kind,
        inlines: block.inlines.map((inline): Inline => ({ ...inline })),
      }
    }),
    comments: doc.comments.map((c) => ({ ...c })),
  }
}
