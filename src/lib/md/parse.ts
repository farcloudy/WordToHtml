/**
 * 类 md 源码 → 文档模型。
 *
 * 这是一套「特定模式」的简化语法，不是通用 Markdown：
 * 一个非空行就是一段，行与行之间不会合并成同一段（公文写作习惯是一行一段，
 * 用空行分隔反而要多敲一次回车）。空行只起分隔作用，本身不产出内容。
 *
 * 块级标记（行首）：
 *   #      文本标题
 *   ##     一级标题      → 自动编号 一、
 *   ###    二级标题      → 自动编号 （一）
 *   ####   三级标题      → 自动编号 1、
 *   @      抬头（正文但取消首行缩进）
 *   >>     落款（正文但右对齐）
 *   -      列表段落
 *   !      列表标题
 *   ---    分节符（单独一行，新起一页且页码重新从 1 开始）
 *
 * 行内标记：
 *   **文字**          加粗
 *   {红|文字}         改色，支持中文色名或 #RRGGBB
 *   {+文字}           插入修订
 *   {-文字}           删除修订
 *   [[文字|批注内容]]  批注，锚定在「文字」上
 *   反斜杠 \ 转义上述所有标记字符
 *
 * 行内标记可以互相嵌套，例如 {红|**重点**}、{+**新增且加粗**}。
 */

import type { BlockKind } from '../spec'
import type { Block, CommentDef, DocModel, Inline, RevMark } from '../types'
import { nextBlockId } from '../types'

export interface ParseOptions {
  /** 修订与批注的作者名，写进 docx 的 w:author */
  author?: string
  /** 时间戳来源，测试里注入固定值以获得确定性输出 */
  now?: () => Date
  /** 修订 id 起始值 */
  startRevisionId?: number
}

const COLOR_NAMES: Record<string, string> = {
  红: 'FF0000',
  黑: '000000',
  蓝: '0000FF',
  绿: '008000',
  紫: '800080',
  橙: 'FF8C00',
  灰: '808080',
  黄: 'BF8F00',
  白: 'FFFFFF',
}

function normalizeColor(token: string): string | undefined {
  const t = token.trim()
  if (/^#?[0-9a-fA-F]{6}$/.test(t)) return t.replace(/^#/, '').toUpperCase()
  return COLOR_NAMES[t]
}

interface InlineContext {
  bold: boolean
  color?: string
  rev?: RevMark
  comments: CommentDef[]
  author: string
  now: () => Date
  nextRevisionId: () => number
  nextCommentId: () => number
}

/** 找到与 index 处字符配对的闭合标记，处理嵌套与转义 */
function findMatchingBrace(src: string, openIndex: number): number {
  let depth = 0
  for (let i = openIndex; i < src.length; i += 1) {
    const c = src[i]
    if (c === '\\') {
      i += 1
      continue
    }
    if (c === '{') depth += 1
    else if (c === '}') {
      depth -= 1
      if (depth === 0) return i
    }
  }
  return -1
}

function findClosing(src: string, from: number, token: string): number {
  let i = from
  while (i < src.length) {
    const hit = src.indexOf(token, i)
    if (hit < 0) return -1
    // 数一下 token 前面连续的反斜杠，奇数表示被转义
    let backslashes = 0
    let j = hit - 1
    while (j >= 0 && src[j] === '\\') {
      backslashes += 1
      j -= 1
    }
    if (backslashes % 2 === 0) return hit
    i = hit + token.length
  }
  return -1
}

/**
 * 找顶层（不在 {} 内、未被转义）的第一个竖线。
 * 指令与批注都用它分隔左右两半，这样 {#FF0000|文字} 里的竖线不会被误判成分隔符。
 */
function findTopLevelBar(src: string): number {
  let depth = 0
  let i = 0
  while (i < src.length) {
    const c = src[i]
    if (c === '\\') {
      i += 2
      continue
    }
    if (c === '{') depth += 1
    else if (c === '}') depth -= 1
    else if (c === '|' && depth === 0) return i
    i += 1
  }
  return -1
}

function makeRev(kind: 'ins' | 'del', ctx: InlineContext): RevMark {
  return {
    kind,
    id: ctx.nextRevisionId(),
    author: ctx.author,
    date: ctx.now().toISOString(),
  }
}

/** 解析一段花括号指令，返回它产出的 inline 序列 */
function parseDirective(inner: string, ctx: InlineContext): Inline[] {
  const head = inner[0]
  if (head === '+') {
    return parseInline(inner.slice(1), { ...ctx, rev: makeRev('ins', ctx) })
  }
  if (head === '-') {
    return parseInline(inner.slice(1), { ...ctx, rev: makeRev('del', ctx) })
  }

  const bar = findTopLevelBar(inner)
  if (bar < 0) return literal(inner)
  const color = normalizeColor(inner.slice(0, bar))
  if (!color) return literal(inner)
  return parseInline(inner.slice(bar + 1), { ...ctx, color })
}

function literal(text: string): Inline[] {
  return text === '' ? [] : [{ t: 'text', text }]
}

function parseInline(src: string, ctx: InlineContext): Inline[] {
  const out: Inline[] = []
  let buffer = ''

  const flush = (): void => {
    if (buffer === '') return
    out.push({
      t: 'text',
      text: buffer,
      ...(ctx.bold ? { bold: true } : {}),
      ...(ctx.color ? { color: ctx.color } : {}),
      ...(ctx.rev ? { rev: ctx.rev } : {}),
    })
    buffer = ''
  }

  let i = 0
  while (i < src.length) {
    const c = src[i]

    if (c === '\\' && i + 1 < src.length) {
      buffer += src[i + 1]
      i += 2
      continue
    }

    if (src.startsWith('**', i)) {
      const close = findClosing(src, i + 2, '**')
      if (close >= 0 && close > i + 2) {
        flush()
        out.push(...parseInline(src.slice(i + 2, close), { ...ctx, bold: true }))
        i = close + 2
        continue
      }
    }

    if (c === '{') {
      const close = findMatchingBrace(src, i)
      if (close > i + 1) {
        flush()
        out.push(...parseDirective(src.slice(i + 1, close), ctx))
        i = close + 1
        continue
      }
    }

    if (src.startsWith('[[', i)) {
      const close = src.indexOf(']]', i + 2)
      if (close > i + 2) {
        const inner = src.slice(i + 2, close)
        const bar = findTopLevelBar(inner)
        if (bar > 0) {
          flush()
          const id = ctx.nextCommentId()
          ctx.comments.push({
            id,
            author: ctx.author,
            date: ctx.now().toISOString(),
            text: inner.slice(bar + 1),
          })
          out.push({ t: 'commentStart', commentId: id })
          out.push(...parseInline(inner.slice(0, bar), ctx))
          out.push({ t: 'commentEnd', commentId: id })
          i = close + 2
          continue
        }
      }
    }

    buffer += c
    i += 1
  }

  flush()
  return out
}

interface BlockRule {
  re: RegExp
  kind: BlockKind
}

// 顺序要紧：长标记必须排在它的前缀之前（#### 先于 ### 先于 ## 先于 #）
const BLOCK_RULES: readonly BlockRule[] = [
  { re: /^####\s+(.*)$/, kind: 'h3' },
  { re: /^###\s+(.*)$/, kind: 'h2' },
  { re: /^##\s+(.*)$/, kind: 'h1' },
  { re: /^#\s+(.*)$/, kind: 'title' },
  { re: /^@\s*(.*)$/, kind: 'salutation' },
  { re: /^>>\s*(.*)$/, kind: 'signature' },
  { re: /^!\s+(.*)$/, kind: 'listTitle' },
  { re: /^-\s+(.*)$/, kind: 'listItem' },
]

const SECTION_BREAK_RE = /^-{3,}\s*$/

export function parseMd(source: string, options: ParseOptions = {}): DocModel {
  const author = options.author ?? '管理员'
  const now = options.now ?? (() => new Date())
  let revisionSeq = options.startRevisionId ?? 0
  let commentSeq = 0

  const comments: CommentDef[] = []
  const blocks: Block[] = []

  const baseCtx = {
    bold: false,
    comments,
    author,
    now,
    nextRevisionId: () => {
      const id = revisionSeq
      revisionSeq += 1
      return id
    },
    nextCommentId: () => {
      const id = commentSeq
      commentSeq += 1
      return id
    },
  }

  for (const rawLine of source.split(/\r?\n/)) {
    const line = rawLine.trimEnd()
    if (line.trim() === '') continue

    if (SECTION_BREAK_RE.test(line)) {
      blocks.push({ t: 'sectionBreak', id: nextBlockId('s'), restartNumbering: true })
      continue
    }

    let matched = false
    for (const rule of BLOCK_RULES) {
      const m = rule.re.exec(line)
      const body = m?.[1]
      if (body === undefined) continue
      blocks.push({
        t: 'textBlock',
        id: nextBlockId(),
        kind: rule.kind,
        inlines: parseInline(body, baseCtx),
      })
      matched = true
      break
    }
    if (matched) continue

    blocks.push({
      t: 'textBlock',
      id: nextBlockId(),
      kind: 'body',
      inlines: parseInline(line, baseCtx),
    })
  }

  return { blocks, comments }
}
