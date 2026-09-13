/**
 * 文档模型。
 *
 * 这是整个组件的真理来源：md 字符串、预览 DOM、docx 文件三者都只经由模型互相转换，
 * 不在彼此之间直接对穿。这样修订、批注这类需要额外元信息的东西才有地方安放。
 */

import type { Align, BlockKind } from './spec'

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
  underline?: boolean
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

/**
 * 软换行，对应 Word 的 `<w:br/>`（单元格里 Shift+Enter 的产物）。
 *
 * **零宽**：它不占字符位 —— plainText、块长度、分页的字符偏移坐标系、
 * sliceInlines/sliceStrict 的区间全都按「它不存在」来算，所以量测与分页算术
 * 一行都不用改。它只影响渲染（多一个行盒）与导出（一个 <w:br/>）。
 */
export interface BreakInline {
  t: 'break'
}

export type Inline = TextInline | CommentStartInline | CommentEndInline | BreakInline

/** 纸张方向。landscape 时纸张宽高互换（换算只在 resolveSections 一处做）。 */
export type PageOrientation = 'portrait' | 'landscape'

/**
 * 一节的设置。字段全部可选：**只存被显式改过的值**，缺省由
 * resolveSectionSettings() 一处解析（预览 / 分页 / 导出 / 上下文工具条同源）。
 *
 * 为什么不把默认值写进模型：默认值是一份真相，写进模型就成了两份；
 * 而且「只写非默认值」正是 md 往返字节稳定的前提（见 md/serialize.ts）。
 */
export interface SectionSettings {
  /** 本节是否显示页码（默认 true）。linkPrevious 为 true 时无意义 */
  pageNumbers?: boolean
  /** 页码与页脚是否沿用前一节（默认 true）。首节没有前节，解析时恒 false */
  linkPrevious?: boolean
  /** 本节页码是否从 1 重新起算（默认 false）。linkPrevious 为 true 时无意义 */
  restartAtOne?: boolean
  /** 纸张方向（默认 portrait） */
  orientation?: PageOrientation
}

/**
 * 分节符。对应 md 里的独立一行 `---`（可带 kwargs）。
 *
 * 它只是**分隔符**：节的具体设置挂在下标对应的 DocModel.sections 上，
 * 不随分节符走 —— 这样首节（没有分节符承载它）与其余各节在形状上完全对称。
 */
export interface SectionBreakBlock {
  t: 'sectionBreak'
  id: string
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

/**
 * 表格行的角色。`unit`（表头行上方的「单位：元」）与 `note`（表尾附注）**天然就是整行一格**，
 * 所以模型里只存 1 个 cell，渲染与导出时再展开成 columnSpan = columns ——「整行合并」
 * 因此不可能与实际格数不一致，也就没有合并字段需要维护。
 */
export type TableRowRole = 'unit' | 'body' | 'note'

/** 垂直对齐三档（逐格覆盖用；Word 的 w:vAlign） */
export type CellVerticalAlign = 'top' | 'middle' | 'bottom'

/**
 * 单元格的对齐覆盖。只有被用户显式改过的维度才存字段 ——
 * 缺省值由 render/html.ts、docx/export.ts 与 emitSelection 三处按同一条规则现算
 * （见 defaultCellAlignH；垂直缺省一律 top）。
 */
export interface TableCellAlign {
  h?: Align
  v?: CellVerticalAlign
}

/**
 * 单元格内容：单段落；软换行（Shift+Enter 的 <w:br/>）作为零宽 inline 存在 inlines 里。
 *
 * `kind` 缺省是 `listItem`（列表段落）—— 老样本与老 docx 因此一个字节都不用变；
 * 显式设成 `listItem` 时 `setCellKind` 会把字段删掉（模型里不存冗余值）。
 */
export interface TableCellModel {
  inlines: Inline[]
  kind?: BlockKind
  align?: TableCellAlign
}

export interface TableRowModel {
  role: TableRowRole
  cells: TableCellModel[]
}

export interface TableBlock {
  t: 'table'
  id: string
  /** 顺序即显示顺序：unit → body… → note；解析时保持 md 里的书写顺序，不强制重排 */
  rows: TableRowModel[]
  /** body 行的最大格数；解析时保证 >= 1 */
  columns: number
  /** 「最小一行 / 最小两行」，即 HeightRule.ATLEAST 的倍数 */
  minLines: 1 | 2
  /** 默认 true（「默认禁止跨页断行」= 行不跨页断开，即 Word 的 w:cantSplit） */
  cantSplit: boolean
}

export type Block = TextBlock | SectionBreakBlock | PageBreakBlock | TableBlock

/**
 * 单元格在编辑层里的「伪块」id。
 *
 * 单元格也挂 data-block-id，于是 edit/dom.ts 的坐标换算（fragmentOf 向上取最近的
 * data-block-id）不用改一行就能把光标落进格子里 —— 代价只是编辑层按 id 找容器。
 */
export function cellId(tableId: string, r: number, c: number): string {
  return `${tableId}.r${r}c${c}`
}

const CELL_ID_RE = /^(.*)\.r(\d+)c(\d+)$/

/** cellId() 的反向解析；不是单元格 id 时返回 null */
export function parseCellId(id: string): { tableId: string; row: number; col: number } | null {
  const m = CELL_ID_RE.exec(id)
  const tableId = m?.[1]
  if (m === null || !tableId) return null
  return { tableId, row: Number(m[2]), col: Number(m[3]) }
}

/**
 * 逐格水平对齐的默认值（没有 `cell.align.h` 覆盖时用它）。
 *
 * unit 行恒右对齐、note 行恒左对齐（这是两个 role 的既有语义，见 6.1）；
 * body 格跟着自己那条样式走 —— 调用方要把样式的 align 传进来
 * （render/html.ts 里 body 格靠类名吃到样式对齐，不需要这个函数，所以它传什么都不会被用到）。
 */
export function defaultCellAlignH(role: TableRowRole, styleAlign: Align): Align {
  if (role === 'unit') return 'right'
  if (role === 'note') return 'left'
  return styleAlign
}

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
  /**
   * 逐节设置，**下标 = 节号**（0 = 首节）。
   *
   * 不变式：`sections.length === 分节符数 + 1`；整篇只有一节且全默认时可以不写这个字段。
   * 只存非默认值：整节全默认时这一项是 `{}`。
   * 解析侧（resolveSections）必须容忍字段缺失 / 数组偏短（手搓模型、旧 md、旧 docx 读回），
   * 缺项一律按全默认处理。
   */
  sections?: SectionSettings[]
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

/** 只统计文字，批注锚点与软换行都不占字符。分页切分用的字符偏移就是基于这个坐标系。 */
export function plainText(block: Block): string {
  if (block.t !== 'textBlock') return ''
  return inlinesText(inlinesOf(block))
}

/** inline 序列里的纯文字（批注锚点、软换行都不占字符） */
export function inlinesText(inlines: readonly Inline[]): string {
  let s = ''
  for (const inline of inlines) {
    if (inline.t === 'text') s += inline.text
  }
  return s
}

/** 单个 textBlock 的 inline；表格格子没有这个（没有「块」的概念） */
function inlinesOf(block: Block): readonly Inline[] {
  return block.t === 'textBlock' ? block.inlines : []
}

/** 「持有一串 inline」的东西：段落与表格格子都是这个形状 */
export interface InlineHolder {
  inlines: Inline[]
}

/**
 * 全篇所有可编辑容器的 inlines（段落 + 表格每个格子），按文档顺序。
 *
 * 修订标记、批注锚点都可能落在格子里，所以全篇扫描（找修订 id 最大值、清批注锚点……
 * 这类事）必须走这一条，只遍历 blocks 里能看到的 textBlock 会漏。
 */
export function allInlineHolders(doc: DocModel): InlineHolder[] {
  const out: InlineHolder[] = []
  for (const block of doc.blocks) {
    if (block.t === 'textBlock') {
      out.push(block)
      continue
    }
    if (block.t !== 'table') continue
    for (const row of block.rows) {
      for (const cell of row.cells) out.push(cell)
    }
  }
  return out
}

/** 取块内所有修订标记，便于统计/校验 */
export function collectRevisions(doc: DocModel): RevMark[] {
  const out: RevMark[] = []
  for (const holder of allInlineHolders(doc)) {
    for (const inline of holder.inlines) {
      if (inline.t === 'text' && inline.rev) out.push(inline.rev)
    }
  }
  return out
}

/**
 * 每条批注锚定的文字（Word 里叫 scope），按 id 索引。
 * 批注可以嵌套，所以用栈而不是单个游标；锚点不成对时丢弃，不抛错。
 * 软换行零宽（不采字），表格格子里的锚点同样算数。
 */
export function commentScopes(doc: DocModel): Map<number, string> {
  const out = new Map<number, string>()
  const open: { id: number; text: string }[] = []

  for (const holder of allInlineHolders(doc)) {
    for (const inline of holder.inlines) {
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
      if (inline.t !== 'text') continue
      for (const entry of open) entry.text += inline.text
    }
  }
  return out
}

/**
 * 按字符区间切出 inline 片段（分页时把跨页的段落切开用）。
 * 批注锚点会跟着它夹住的那段文字一起被切走：起点落在 [from,to) 内才保留，
 * 终点同理，这样跨页后批注仍然锚定在同一段文字上。
 *
 * 软换行是零宽的，只能按「落在哪一片里」来分：片首（from>0 的续排片）的那一枚丢掉 ——
 * 页边界本来就已经把这一行断开了，再多渲一个 <br> 会凭空多出一个行盒，
 * 而分页算术里没有为它记账（量到的行盒与渲出来的必须一致，否则页面会溢出）。
 */
export function sliceInlines(inlines: readonly Inline[], from: number, to: number): Inline[] {
  const out: Inline[] = []
  let cursor = 0
  for (const inline of inlines) {
    if (inline.t === 'break') {
      if (cursor < to && (from === 0 || cursor > from)) out.push(inline)
      continue
    }
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
