/**
 * 文档模型 → HTML 片段。
 *
 * 预览组件用 v-html、测量器用 innerHTML，两边走同一个函数，
 * 这样「量到的」和「看到的」一定是同一套 DOM 结构。
 */

import type { Inline } from '../types'

export function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}

/**
 * 渲染一段 inline 序列。
 *
 * prefix 是自动编号（如「一、」），跟在段首、随段落样式一起继承加粗。
 * 批注用 <span> 包住被锚定的文字，批注内容本身不进正文。
 */
export function renderInlinesHtml(inlines: readonly Inline[], prefix = ''): string {
  let out = escapeHtml(prefix)
  const openComments: number[] = []

  for (const inline of inlines) {
    if (inline.t === 'commentStart') {
      out += `<span class="wtp-comment" data-comment="${inline.commentId}">`
      openComments.push(inline.commentId)
      continue
    }
    if (inline.t === 'commentEnd') {
      const at = openComments.lastIndexOf(inline.commentId)
      if (at >= 0) {
        out += '</span>'
        openComments.splice(at, 1)
      }
      continue
    }

    let inner = escapeHtml(inline.text)
    if (inline.rev) inner = `<span class="wtp-rev-${inline.rev.kind}">${inner}</span>`
    if (inline.color) inner = `<span style="color:#${inline.color}">${inner}</span>`
    if (inline.bold) inner = `<b>${inner}</b>`
    out += inner
  }

  // 批注锚点若不成对（模型异常），在这里补上闭合，避免把整篇都吞进高亮里
  for (let i = openComments.length - 1; i >= 0; i -= 1) out += '</span>'
  return out
}
