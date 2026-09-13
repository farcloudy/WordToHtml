/**
 * 导航窗格的大纲（纯函数）。
 *
 * 只收 h1 / h2 / h3：title 不进大纲（与 Word 的导航窗格一致），h4 在本项目里
 * 就是 h3 样式、不是单独的块型。
 */

import type { DocModel } from '../types'
import { plainText } from '../types'

export interface OutlineEntry {
  blockId: string
  level: 1 | 2 | 3
  /** 自动编号（如「一、」），没有编号的块为空串 */
  prefix: string
  /** 块内纯文字，不含编号 */
  text: string
}

const LEVEL_OF: Record<string, 1 | 2 | 3 | undefined> = { h1: 1, h2: 2, h3: 3 }

export function buildOutline(doc: DocModel, numbering: Map<string, string>): OutlineEntry[] {
  const out: OutlineEntry[] = []
  for (const block of doc.blocks) {
    if (block.t !== 'textBlock') continue
    const level = LEVEL_OF[block.kind]
    if (level === undefined) continue
    out.push({
      blockId: block.id,
      level,
      prefix: numbering.get(block.id) ?? '',
      text: plainText(block),
    })
  }
  return out
}

/**
 * 大纲的指纹：只有它变了才需要对外发事件。
 * 打字时块 id 不变、正文每敲一个字都变，所以指纹会把「改标题文字」也算成变化；
 * 真正要挡住的是「其它段落打字把左栏带着重渲染」——那种情况指纹完全不动。
 */
export function outlineSignature(entries: readonly OutlineEntry[]): string {
  return entries
    .map((e) => `${e.blockId}\u0000${e.level}\u0000${e.prefix}\u0000${e.text}`)
    .join('\n')
}
