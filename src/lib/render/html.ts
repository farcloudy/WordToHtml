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

import type { Align } from '../spec'
import type { Inline, TableBlock, TableCellModel, TableRowModel } from '../types'
import { cellId, defaultCellAlignH, inlinesText } from '../types'

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
  if (out === '') return '<br>'

  /*
   * 段尾（格尾）那枚软换行之后再补一枚**占位** `<br>`。两件事都指着它：
   *
   * ① 行数要与 Word 一致 —— 「文字 + 行尾软换行」在 Word 里占**两行**（第二行是空的），
   *    而浏览器不给尾随的 `<br>` 单独开行盒（实测 `abc<br>` 量出来还是 1 行）。
   * ② 给 Chrome 一个能落笔的位置 —— 光标停在那枚尾随 `<br>` 之后时，敲进来的字会被插到
   *    它**之前**（回到上一行）；补一枚占位 `<br>` 才有第二行可落，读回也不会把模型里那枚
   *    软换行复制成两枚（2026-09-16 实测，见 issues/20260916-1 第 1 条与 PLAN 8.3）。
   *
   * 占位 `<br>` 不带 `wtp-br` 类，`edit/dom.ts` 的 readInlines 一律忽略裸 `<br>`，
   * 所以它只是排版与手感的脚手架，不进模型。
   */
  return inlines[inlines.length - 1]?.t === 'break' ? `${out}<br>` : out
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
 *   · 格内那层 div 挂 `wtp-cell wtp-<kind>` + `data-block-id = cellId(...)` +
 *     `data-from="0"` / `data-to="<格内文字长度>"`。`wtp-cell` 是给验收脚本与 CSS 用的
 *     稳定钩子，`wtp-<kind>` 才是那条样式（缺省 `listItem`，与 W4a 起的默认一致）。
 *     这样 `edit/dom.ts` 的坐标换算一行都不用改：格内光标自然落到这一层上。
 *   · `unit` / `note` 行天然整行一格（`colspan = columns`、无边框）。它们的角色默认对齐
 *     （unit 右、note 左）写在格内 div 的行内 style 上；逐格覆盖也走同一条路。
 *     垂直对齐只能写在 `<td>` 上（格内 div 上的 vertical-align 无效）。
 *   · `<td>` 挂 `wtp-td` + `wtp-td-<kind>`（plain 行另有 `wtp-td-plain`）+ 恒定的
 *     `data-cell-id`（与格内 div 的 `data-block-id` 同一个值），
 *     行高的**最小值**由 CSS 按这个类名写在 `td` 的 `height` 上（不是格内 div 的 min-height）
 *     —— 写在内层 div 上会把 div 本身撑成两行高，`vertical-align` 就再也挪不动那一行字了
 *     （div 已经填满整个格子，对齐的是 div 而不是 div 里的字）。见 css.ts 的表格一段。
 *     `data-cell-id` 是给**整格复选**的高亮用的：它只由模型决定（与选中态无关），
 *     所以量测缓存签名、逐块量测对账都不受影响；选中高亮由组件在 DOM 上按它挂类名
 *     （见 css.ts 的 CELL_SELECTION_CLASS）。
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

/** 对齐值 → CSS 的 text-align（规格表里的 both 在 CSS 里叫 justify） */
function cssTextAlign(a: Align): string {
  return a === 'both' ? 'justify' : a
}

/** 格内那一层 div：`r` 是 rows 数组下标，与 cellId 同一套 */
function tableCellHtml(
  tableId: string,
  r: number,
  c: number,
  cell: TableCellModel,
  divStyle = '',
): string {
  const length = inlinesText(cell.inlines).length
  const kind = cell.kind ?? 'listItem'
  const style = divStyle === '' ? '' : ` style="${divStyle}"`
  return (
    `<div class="wtp-cell wtp-${kind}" data-block-id="${cellId(tableId, r, c)}" ` +
    `data-from="0" data-to="${length}"${style}>${renderInlinesHtml(cell.inlines)}</div>`
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
    // 水平默认值来自 `wtp-<kind>` 那条样式（body 格跟着自己的样式），只有覆盖才写行内；
    // 垂直对齐的默认是 top（全局 CSS 已是 top），只有覆盖才写 <td> 的行内样式。
    const divStyle = cell.align?.h ? `text-align: ${cssTextAlign(cell.align.h)}` : ''
    const tdStyle = cell.align?.v ? ` style="vertical-align:${cell.align.v}"` : ''
    const tdClass = `wtp-td wtp-td-${cell.kind ?? 'listItem'}`
    cells.push(
      `<td class="${tdClass}" data-cell-id="${cellId(block.id, r, c)}"${tdStyle}>` +
        `${tableCellHtml(block.id, r, c, cell, divStyle)}</td>`,
    )
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
  // unit 右对齐、note 左对齐是角色默认（不是覆盖），所以这里也照写行内 ——
  // 格内的 `wtp-listItem` 类把它默认成 justify，不写就丢了对齐。
  const h = cell.align?.h ?? defaultCellAlignH(row.role, 'both')
  const v = cell.align?.v ?? 'top'
  const divStyle = `text-align: ${cssTextAlign(h)}`
  const tdClass = `wtp-td wtp-td-plain wtp-td-${cell.kind ?? 'listItem'}`
  return (
    `<tr class="wtp-tr-plain"><td class="${tdClass}" data-cell-id="${cellId(block.id, r, 0)}" ` +
    `colspan="${columns}" style="vertical-align:${v}">` +
    `${tableCellHtml(block.id, r, 0, cell, divStyle)}</td></tr>`
  )
}
