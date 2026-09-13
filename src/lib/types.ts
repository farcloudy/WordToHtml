/**
 * 文档模型。
 *
 * 这是整个组件的真理来源：md 字符串、预览 DOM、docx 文件三者都只经由模型互相转换，
 * 不在彼此之间直接对穿。这样修订、批注这类需要额外元信息的东西才有地方安放。
 */

import type { BlockKind } from './spec'

/** 修订标记。Word 的 w:ins / w:del 需要 id + author + date 三个属性，缺一会显示异常。 */
export interface RevMark {
  kind: 'ins' | 'del'
  id: number
  author: string
  /** ISO 8601 */
  date: string
}

export interface TextInline {
  t: 'text'
  text: string
  bold?: boolean
  /** 十六进制颜色，不带 #，例如 'FF0000' */
  color?: string
  rev?: RevMark
}

/** 批注锚点。成对出现，之间夹着被批注的文字。 */
export interface CommentStartInline {
  t: 'commentStart'
  commentId: number
}

export interface CommentEndInline {
  t: 'commentEnd'
  commentId: number
}

export type Inline = TextInline | CommentStartInline | CommentEndInline

/**
 * 分节符。对应 md 里的独立一行 `---`。
 * 分节在 Word 里意味着：新的一页 + 独立的页眉页脚（页码可以重新从 1 开始）。
 */
export interface SectionBreakBlock {
  t: 'sectionBreak'
  id: string
  /** 本节页码是否从 1 重新开始 */
  restartNumbering: boolean
}

/**
 * 分页符。对应 md 里的独立一行 `===`。
 * 与分节符的区别：只强制换页，不新开一节（页码连续、页眉页脚照旧）。
 */
export interface PageBreakBlock {
  t: 'pageBreak'
  id: string
}

export interface TextBlock {
  t: 'textBlock'
  id: string
  kind: BlockKind
  inlines: Inline[]
}

export type Block = TextBlock | SectionBreakBlock | PageBreakBlock

export interface CommentDef {
  id: number
  author: string
  /** ISO 8601 */
  date: string
  text: string
  /** 回复某条批注时指向父批注 id */
  parentId?: number
  /** 整条会话标记为已解决 */
  resolved?: boolean
}

export interface DocModel {
  blocks: Block[]
  comments: CommentDef[]
}

export const emptyDoc = (): DocModel => ({ blocks: [], comments: [] })

/* -------------------------------------------------------------------------- */
/* 纯函数辅助                                                                  */
/* -------------------------------------------------------------------------- */

let idSeq = 0

/** 生成模型内部用的 block id。只在内存里用，不写进 docx。 */
export function nextBlockId(prefix = 'b'): string {
  idSeq += 1
  return `${prefix}${idSeq}`
}

/** 只统计文字，批注锚点不占字符。分页切分用的字符偏移就是基于这个坐标系。 */
export function plainText(block: Block): string {
  if (block.t !== 'textBlock') return ''
  let s = ''
  for (const inline of block.inlines) {
    if (inline.t === 'text') s += inline.text
  }
  return s
}

/** 取块内所有修订标记，便于统计/校验 */
export function collectRevisions(doc: DocModel): RevMark[] {
  const out: RevMark[] = []
  for (const block of doc.blocks) {
    if (block.t !== 'textBlock') continue
    for (const inline of block.inlines) {
      if (inline.t === 'text' && inline.rev) out.push(inline.rev)
    }
  }
  return out
}

/**
 * 每条批注锚定的文字（Word 里叫 scope），按 id 索引。
 * 批注可以嵌套，所以用栈而不是单个游标；锚点不成对时丢弃，不抛错。
 */
export function commentScopes(doc: DocModel): Map<number, string> {
  const out = new Map<number, string>()
  const open: { id: number; text: string }[] = []

  for (const block of doc.blocks) {
    if (block.t !== 'textBlock') continue
    for (const inline of block.inlines) {
      if (inline.t === 'commentStart') {
        open.push({ id: inline.commentId, text: '' })
        continue
      }
      if (inline.t === 'commentEnd') {
        const at = open.map((o) => o.id).lastIndexOf(inline.commentId)
        if (at >= 0) {
          const entry = open.splice(at, 1)[0]
          if (entry) out.set(entry.id, entry.text)
        }
        continue
      }
      for (const entry of open) entry.text += inline.text
    }
  }
  return out
}

/**
 * 按字符区间切出 inline 片段（分页时把跨页的段落切开用）。
 * 批注锚点会跟着它夹住的那段文字一起被切走：起点落在 [from,to) 内才保留，
 * 终点同理，这样跨页后批注仍然锚定在同一段文字上。
 */
export function sliceInlines(inlines: readonly Inline[], from: number, to: number): Inline[] {
  const out: Inline[] = []
  let cursor = 0
  for (const inline of inlines) {
    if (inline.t !== 'text') {
      out.push(inline)
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
