/**
 * 文档模型 → HTML 片段。
 *
 * 预览组件用 v-html、测量器用 innerHTML，两边走同一个函数，
 * 这样「量到的」和「看到的」一定是同一套 DOM 结构。
 *
 * 编辑层也依赖这一套结构：读回 DOM 时靠 .wtp-num / .wtp-comment /
 * .wtp-rev-* 这些类名、data-* 属性与 <b> / <u> 这两个行内标签还原模型，
 * 所以它们的形态是接口的一部分，改类名、去掉 data-* 或改标签名会同时打断编辑与量测。
 */

import type { Inline, TableBlock, TableRowModel } from '../types'
import { cellId, inlinesText } from '../types'

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
    if (inline.t === 'break') {
      // 真软换行。带一个类名是为了与「空段落占位」的那枚裸 <br> 区分开 ——
      // 读回模型时（edit/dom.ts 的 readInlines）只认带 wtp-br 的，裸 <br> 一律忽略。
      out += `<br class="wtp-br">`
      continue
    }
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
    if (inline.underline) inner = `<u>${inner}</u>`
    out += inner
  }

  // 批注锚点若不成对（模型异常），在这里补上闭合，避免把整篇都吞进高亮里
  for (let i = openComments.length - 1; i >= 0; i -= 1) out += '</span>'

  // 空段落也要占一行 —— Word 里空段落就是一个行高的占位。不给这一口 <br>，
  // 块高会量成 0：既与 Word 的分页对不上，也没法把插入符放进一个没有行盒的块里。
  // 读回模型时 <br> 会被忽略，所以它不会变成内容。（带 wtp-br 类的那枚是真软换行，不算占位。）
  return out === '' ? '<br>' : out
}

/**
 * 表格块 → HTML 片段（**量测与预览共用这一个函数**）。
 *
 * 量测的可信前提是「量到的就是看到的」，所以预览与探针必须渲出同一套 DOM。
 *
 * 类名与 data-* 是接口的一部分，改动会同时打断编辑与量测：
 *   · 外层 `<table>` 挂 `wtp-table`，**绝不能挂 data-block-id** ——
 *     `fragmentOf` 向上取最近的 `data-block-id`，而 `retagFragments` / `syncPlain`
 *     用 `querySelectorAll('[data-block-id]')` 找片段；外层挂了就会被当成一个片段，
 *     整张表会被当成一段文字读回模型。外层改用 `data-table-id` + `data-row-from/to`。
 *   · 格内那层 div 复用 `wtp-listItem`（列表段落样式；W1 已把它的首行缩进改成 0，
 *     正是为了让格内文字不用再加缩进豁免），并挂
 *     `data-block-id = cellId(...)` + `data-from="0"` / `data-to="<格内文字长度>"`。
 *     这样 `edit/dom.ts` 的坐标换算一行都不用改：格内光标自然落到这一层上。
 *   · `unit` / `note` 行天然整行一格（`colspan = columns`、无边框）。本项目的段落对齐
 *     是样式级的、没有逐段覆盖，所以这两行的对齐只能写在格内 div 的行内 style 上。
 */
export function renderTableFragment(block: TableBlock, rowFrom: number, rowTo: number): string {
  const columns = Math.max(1, block.columns)
  const from = Math.max(0, rowFrom)
  const to = Math.min(block.rows.length, Math.max(from, rowTo))
  const rows: string[] = []
  for (let r = from; r < to; r += 1) {
    const row = block.rows[r]
    if (!row) continue
    rows.push(row.role === 'body' ? tableBodyRow(block, row, r, columns) : tablePlainRow(block, row, r, columns))
  }
  const minClass = block.minLines === 2 ? 'wtp-table-min2' : 'wtp-table-min1'
  return `<table class="wtp-table ${minClass}"><tbody>${rows.join('')}</tbody></table>`
}

/** 格内那一层 div：复用列表段落样式 + 格内坐标（`r` 是 rows 数组下标，与 cellId 同一套） */
function tableCellHtml(
  tableId: string,
  r: number,
  c: number,
  inlines: readonly Inline[],
  inlineStyle = '',
): string {
  const length = inlinesText(inlines).length
  const style = inlineStyle === '' ? '' : ` style="${inlineStyle}"`
  return (
    `<div class="wtp-listItem" data-block-id="${cellId(tableId, r, c)}" ` +
    `data-from="0" data-to="${length}"${style}>${renderInlinesHtml(inlines)}</div>`
  )
}

function tableBodyRow(
  block: TableBlock,
  row: TableRowModel,
  r: number,
  columns: number,
): string {
  const cells: string[] = []
  for (let c = 0; c < columns; c += 1) {
    // 缺格补空：Word 的表格必须是矩形；模型仍按 md 原样存（少一格的书写方式不该被解析改写）
    const cell = row.cells[c] ?? { inlines: [] }
    cells.push(`<td>${tableCellHtml(block.id, r, c, cell.inlines)}</td>`)
  }
  return `<tr>${cells.join('')}</tr>`
}

function tablePlainRow(
  block: TableBlock,
  row: TableRowModel,
  r: number,
  columns: number,
): string {
  const cell = row.cells[0] ?? { inlines: [] }
  const unit = row.role === 'unit'
  // unit 右对齐、note 左对齐且顶端对齐（与 docx 导出侧一致；附注行的对齐靠 renderTableFragment 的行内 style）
  const tdStyle = unit ? '' : ' style="vertical-align:top"'
  const divStyle = unit ? 'text-align: right' : 'text-align: left'
  return (
    `<tr class="wtp-tr-plain"><td class="wtp-td-plain" colspan="${columns}"${tdStyle}>` +
    `${tableCellHtml(block.id, r, 0, cell.inlines, divStyle)}</td></tr>`
  )
}
