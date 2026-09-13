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
import type { Block, CommentDef, DocModel, Inline } from '../types'

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
    return { t: 'textBlock', kind: block.kind, inlines: block.inlines }
  })
}
