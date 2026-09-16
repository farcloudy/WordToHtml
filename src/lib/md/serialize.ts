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
import { cellParagraphs } from '../edit/table'
import type {
  Block,
  CommentDef,
  DocModel,
  EditorSettings,
  Inline,
  SectionSettings,
  TableBlock,
  TableCellModel,
} from '../types'
import { headerRowCount } from '../types'

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
 * 单元格内容 → md。**段间写 `{p}`**（格内多段落，`| 甲{p}乙 |` 一格两段）；
 * 格内样式与两组对齐写成格首指令 `{@<kind>,<h>,<v>|正文}`，
 * **token 顺序固定 kind → h → v**、只写非默认值，这样「模型 → md → 模型 → md」字节稳定；
 * 三个都没值时整条指令都不写（老样本因此一个字节不变）。
 *
 * 正文里真写 `{p}` 字面量的可逆性由 escapeText 保证（`{` → `\{`），与 `{br}` 同一条路：
 * 写出去是 `\{p\}`、读回来是字面量 `{p}`，不会越滚越多、也不会被当成段落标记。
 */
function cellText(cell: TableCellModel, comments: Map<number, CommentDef>): string {
  const inner = cellParagraphs(cell)
    .map((para) => serializeInlines(para.inlines, comments))
    .join('{p}')
  const tokens: string[] = []
  if (cell.kind !== undefined && cell.kind !== 'listItem') tokens.push(cell.kind)
  if (cell.align?.h) tokens.push(cell.align.h)
  if (cell.align?.v) tokens.push(cell.align.v)
  return tokens.length === 0 ? inner : `{@${tokens.join(',')}|${inner}}`
}

/**
 * 表格块 → md 围栏。kwarg 只写非默认值（minLines 默认 1、headerRows 默认 0、cantSplit 默认 true），
 * 顺序固定 minLines → headerRows → cantSplit —— 这样「模型 → md → 模型 → md」字节稳定。
 */
function tableLines(block: TableBlock, comments: Map<number, CommentDef>): string[] {
  let fence = ':::table'
  if (block.minLines !== 1) fence += ` minLines=${block.minLines}`
  const header = headerRowCount(block)
  if (header > 0) fence += ` headerRows=${header}`
  if (!block.cantSplit) fence += ' cantSplit=no'

  const out = [fence]
  for (const row of block.rows) {
    if (row.role === 'body') {
      out.push('| ' + row.cells.map((cell) => cellText(cell, comments)).join(' | ') + ' |')
      continue
    }
    const first = row.cells[0]
    const text = first ? cellText(first, comments) : ''
    out.push(row.role === 'unit' ? `> ${text}` : `< ${text}`)
  }
  out.push(':::')
  return out
}

/**
 * 一节的设置 → kwargs 串（**只写非默认值**，token 顺序固定 link → numbers → restart → orientation）。
 *
 * 顺序固定 + 只写非默认值 ⇒「模型 → md → 模型 → md」字节稳定；
 * `isFirst` 时首节没有前节，link 一律不写（它恒无意义，写了也解析不回来）。
 */
function sectionKwargs(raw: SectionSettings | undefined, isFirst: boolean): string {
  if (!raw) return ''
  const tokens: string[] = []
  if (!isFirst && raw.linkPrevious === false) tokens.push('link=off')
  if (raw.pageNumbers === false) tokens.push('numbers=off')
  if (raw.restartAtOne === true) tokens.push('restart=on')
  if (raw.orientation === 'landscape') tokens.push('orientation=landscape')
  return tokens.join(' ')
}

/**
 * 编辑器开关 → `::editor` 的 kwargs（**只写非默认值**，token 顺序固定
 * trackChanges 在前、nav 在后）。全默认时返回空串 —— 这一行整个不写。
 *
 * 顺序固定 + 只写非默认值 ⇒「模型 → md → 模型 → md」字节稳定。
 */
function editorKwargs(editor: EditorSettings | undefined): string {
  if (!editor) return ''
  const tokens: string[] = []
  if (editor.trackChanges === true) tokens.push('trackChanges=on')
  if (editor.nav === false) tokens.push('nav=off')
  return tokens.join(' ')
}

export function toMd(doc: DocModel): string {
  const comments = new Map(doc.comments.map((c) => [c.id, c]))
  const lines: string[] = []
  const sections = doc.sections ?? []

  // 编辑器开关写在文档最前面（它是整个文档的，不属于任何一节）；指令区里
  // `::editor` 在 `::section` 之前，解析侧两者谁先谁后都认
  const editorLine = editorKwargs(doc.editor)
  if (editorLine !== '') lines.push(`::editor ${editorLine}`)

  // 首节写在文档第一行（只有非默认时才写），其余各节跟着各自的 `---` 走
  const head = sectionKwargs(sections[0], true)
  if (head !== '') lines.push(`::section ${head}`)

  // 第 k 个分节符开启第 k+1 节
  let sectionIndex = 0
  for (const block of doc.blocks) {
    if (block.t === 'sectionBreak') {
      sectionIndex += 1
      const kwargs = sectionKwargs(sections[sectionIndex], false)
      lines.push(kwargs === '' ? '---' : `--- ${kwargs}`)
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
      return { t: 'sectionBreak' }
    }
    if (block.t === 'pageBreak') return { t: 'pageBreak' }
    if (block.t === 'table') {
      return {
        t: 'table',
        columns: block.columns,
        minLines: block.minLines,
        cantSplit: block.cantSplit,
        // 归一后的有效值（缺省 / 越界都算 0）—— 与 md 侧「只写非默认值」同一套口径，
        // 否则「headerRows: 0」与「不落字段」会被判成往返不一致
        ...(headerRowCount(block) > 0 ? { headerRows: headerRowCount(block) } : {}),
        rows: block.rows.map((row) => ({
          role: row.role,
          cells: row.cells.map((cell) => ({
            paragraphs: cellParagraphs(cell).map((para) => ({ inlines: para.inlines })),
            ...(cell.kind !== undefined ? { kind: cell.kind } : {}),
            ...(cell.align !== undefined ? { align: cell.align } : {}),
          })),
        })),
      }
    }
    return { t: 'textBlock', kind: block.kind, inlines: block.inlines }
  })
}
