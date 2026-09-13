/**
 * 文档模型 → HTML 片段。
 *
 * 预览组件用 v-html、测量器用 innerHTML，两边走同一个函数，
 * 这样「量到的」和「看到的」一定是同一套 DOM 结构。
 *
 * 编辑层也依赖这一套结构：读回 DOM 时靠 .wtp-num / .wtp-comment /
 * .wtp-rev-* 这些类名与 data-* 属性还原模型，所以它们的形态是接口的一部分，
 * 改类名或去掉 data-* 会同时打断编辑与量测。
 */

import type { Inline } from '../types'

export function escapeHtml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function escapeAttr(text: string): string {
  return escapeHtml(text).replace(/"/g, '&quot;')
}

/**
 * 渲染一段 inline 序列。
 *
 * prefix 是自动编号（如「一、」），跟在段首、随段落样式一起继承加粗。
 * 它包在 contenteditable=false 的 span 里：编号是自动生成的，不该被用户改到，
 * 读回模型时也直接跳过（编号不是模型文字的一部分）。
 *
 * 批注用 <span> 包住被锚定的文字，批注内容本身不进正文 —— 它由调用方
 * 从模型的 comments 里取，渲染在侧栏。
 *
 * activeCommentId 只额外加一个类名（改背景色），不改变盒模型，
 * 因此量测仍可复用同一个函数。
 */
export function renderInlinesHtml(
  inlines: readonly Inline[],
  prefix = '',
  activeCommentId?: number | null,
): string {
  let out = prefix
    ? `<span class="wtp-num" contenteditable="false">${escapeHtml(prefix)}</span>`
    : ''
  const openComments: number[] = []

  for (const inline of inlines) {
    if (inline.t === 'commentStart') {
      const active = activeCommentId != null && activeCommentId === inline.commentId
      out += `<span class="wtp-comment${active ? ' wtp-comment-active' : ''}" data-comment="${inline.commentId}">`
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
    if (inline.rev) {
      // 修订的作者/时间/编号在 DOM 里没有别的地方可取，读回模型时全靠这几个属性
      const rev = inline.rev
      inner =
        `<span class="wtp-rev-${rev.kind}" data-rev="${rev.kind}" data-rev-id="${rev.id}" ` +
        `data-rev-author="${escapeAttr(rev.author)}" data-rev-date="${escapeAttr(rev.date)}">` +
        `${inner}</span>`
    }
    if (inline.color) inner = `<span style="color:#${inline.color}">${inner}</span>`
    if (inline.bold) inner = `<b>${inner}</b>`
    out += inner
  }

  // 批注锚点若不成对（模型异常），在这里补上闭合，避免把整篇都吞进高亮里
  for (let i = openComments.length - 1; i >= 0; i -= 1) out += '</span>'

  // 空段落也要占一行 —— Word 里空段落就是一个行高的占位。不给这一口 <br>，
  // 块高会量成 0：既与 Word 的分页对不上，也没法把插入符放进一个没有行盒的块里。
  // 读回模型时 <br> 会被忽略，所以它不会变成内容。
  return out === '' ? '<br>' : out
}
