/**
 * 文档模型 → 类 md 源码。
 *
 * 用途：把编辑结果落回文本形态（存档、喂给后端、人工比对），并支撑
 * 「模型 → md → 模型」的往返一致性测试。
 *
 * 有意为之的有损之处（不假装无损）：
 * - 修订的作者与时间戳不写进 md，语法里没有这个位置，只在模型和 docx 里保留；
 * - 批注的回复关系（parentId）与已解决状态同样只在模型里保留。
 */

import type { BlockKind } from '../spec'
import type { Block, CommentDef, DocModel, Inline, TableBlock } from '../types'

const PREFIX: Record<BlockKind, string> = {
  title: '# ',
  h1: '## ',
  h2: '### ',
  h3: '#### ',
  salutation: '@ ',
  signature: '>> ',
  attachment: '% ',
  listTitle: '! ',
  listItem: '- ',
  body: '',
}

/** 转义所有会与语法冲突的字符，保证往返可解析回同一份文本 */
function escapeText(text: string): string {
  return text
    .replace(/\\/g, '\\\\')
    .replace(/\|/g, '\\|')
    .replace(/\{/g, '\\{')
    .replace(/\}/g, '\\}')
    .replace(/\*\*/g, '\\*\\*')
    .replace(/__/g, '\\_\\_')
    .replace(/\[\[/g, '\\[[')
}

function serializeInlines(
  inlines: readonly Inline[],
  comments: Map<number, CommentDef>,
): string {
  let out = ''
  for (const inline of inlines) {
    if (inline.t === 'break') {
      // 软换行是零宽标记，不参与 bold/underline/color/rev 的包裹：它没有属于自己的格式，
      // 前后两段文字各自带着自己的格式（readInlines 读回来也是这样）。
      out += '{br}'
      continue
    }
    if (inline.t === 'commentStart') {
      out += '[['
      continue
    }
    if (inline.t === 'commentEnd') {
      const def = comments.get(inline.commentId)
      out += `|${def ? escapeText(def.text) : ''}]]`
      continue
    }

    let s = escapeText(inline.text)
    if (inline.bold) s = `**${s}**`
    if (inline.underline) s = `__${s}__`
    if (inline.color) s = `{#${inline.color}|${s}}`
    if (inline.rev) s = inline.rev.kind === 'ins' ? `{+${s}}` : `{-${s}}`
    out += s
  }
  return out
}

/**
 * 表格块 → md 围栏。kwarg 只写非默认值（minLines 默认 1、cantSplit 默认 true），
 * 顺序固定 minLines 在前、cantSplit 在后 —— 这样「模型 → md → 模型 → md」字节稳定。
 */
function tableLines(block: TableBlock, comments: Map<number, CommentDef>): string[] {
  let fence = ':::table'
  if (block.minLines !== 1) fence += ` minLines=${block.minLines}`
  if (!block.cantSplit) fence += ' cantSplit=no'

  const out = [fence]
  for (const row of block.rows) {
    if (row.role === 'body') {
      out.push(
        '| ' +
          row.cells.map((cell) => serializeInlines(cell.inlines, comments)).join(' | ') +
          ' |',
      )
      continue
    }
    const first = row.cells[0]
    const text = first ? serializeInlines(first.inlines, comments) : ''
    out.push(row.role === 'unit' ? `> ${text}` : `< ${text}`)
  }
  out.push(':::')
  return out
}

export function toMd(doc: DocModel): string {
  const comments = new Map(doc.comments.map((c) => [c.id, c]))
  const lines: string[] = []

  for (const block of doc.blocks) {
    if (block.t === 'sectionBreak') {
      lines.push('---')
      continue
    }
    if (block.t === 'pageBreak') {
      lines.push('===')
      continue
    }
    if (block.t === 'table') {
      lines.push(...tableLines(block, comments))
      continue
    }
    lines.push(PREFIX[block.kind] + serializeInlines(block.inlines, comments))
  }

  return lines.join('\n')
}

/** 便于测试与调试：把块序列归一化成可比较的形状（丢掉内存 id） */
export function normalizeBlocks(doc: DocModel): unknown[] {
  return doc.blocks.map((block: Block) => {
    if (block.t === 'sectionBreak') {
      return { t: 'sectionBreak', restartNumbering: block.restartNumbering }
    }
    if (block.t === 'pageBreak') return { t: 'pageBreak' }
    if (block.t === 'table') {
      return {
        t: 'table',
        columns: block.columns,
        minLines: block.minLines,
        cantSplit: block.cantSplit,
        rows: block.rows.map((row) => ({
          role: row.role,
          cells: row.cells.map((cell) => ({ inlines: cell.inlines })),
        })),
      }
    }
    return { t: 'textBlock', kind: block.kind, inlines: block.inlines }
  })
}
