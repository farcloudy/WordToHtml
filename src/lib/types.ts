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
 * 软换行，对应 Word 的 `<w:br/>`（Shift+Enter 的产物：正文段落与表格格子同一条路）。
 *
 * **零宽**：它不占字符位 —— plainText、块长度、分页的字符偏移坐标系、
 * sliceInlines/sliceStrict 的区间全都按「它不存在」来算，所以量测与分页算术
 * 一行都不用改。它只影响渲染（多一个行盒）与导出（一个 <w:br/>）。
 *
 * 尾随的那一枚（段落/格子以换行结尾）在渲染时会再补一个占位 `<br>`：
 * 浏览器不给尾随 `<br>` 单独开行盒，不补就与 Word 差一行，且敲进来的字会跑回上一行。
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
 * 单元格内容：**多段落**（Word 的单元格本来就放得下多个 `w:p`，Enter 在格内新起一段）。
 *
 * 每段与段落同一个形状（`InlineHolder`），于是编辑层「找容器 → 直接改它的 inlines」那条路
 * 一行都不用改：段落下标进 id（`cellId(...)` 恒指第 0 段，第 N 段是 `` `${cellId(...)}.p${N}` ``）。
 * 软换行（Shift+Enter 的 `<w:br/>`）仍是零宽 inline，落在某一段的 inlines 里。
 *
 * 不变式：`paragraphs.length >= 1`（解析、`normalizeTable`、`cloneDoc` 都保证）。
 *
 * `kind` / `align` 是**格子级**的（Word 的 cell 级属性也只有垂直对齐与宽度），格内各段共用；
 * `kind` 缺省是 `listItem`（列表段落）—— 老样本与老 docx 因此一个字节都不用变；
 * 显式设成 `listItem` 时 `setCellKind` 会把字段删掉（模型里不存冗余值）。
 */
export interface TableCellModel {
  paragraphs: InlineHolder[]
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
  /**
   * 「最小一行 / 最小两行」，即 HeightRule.ATLEAST 的倍数。
   * **只管正文行**：表头行（unit）与附注行（note）恒为「最小一行」，不随它变
   * （见 render/css.ts 与 docx/export.ts 的同一条规则）。
   */
  minLines: 1 | 2
  /** 默认 true（「默认禁止跨页断行」= 行不跨页断开，即 Word 的 w:cantSplit） */
  cantSplit: boolean
  /**
   * 前 N 行是**标题行**，在每一个续页的顶端重复出现（Word 的 w:tblHeader）。
   * 缺省 0 且不落字段（与 minLines / cantSplit 同一套「默认值不落模型」的约定）；
   * 有效区间是 `rows[0 .. headerRows)`，归一化时夹到 `[0, rows.length]`（见 normalizeTable）。
   *
   * 为什么是表格级一个数字而不是行级标记：OOXML 的 w:tblHeader 语义是「这一行**之上**
   * 每一行也都是标题行时才重复」，标在第 3 行而前两行没标等于没标 —— 表格级没有死状态。
   */
  headerRows?: number
}

export type Block = TextBlock | SectionBreakBlock | PageBreakBlock | TableBlock

/**
 * 有效的前导标题行数（区间是 `rows[0 .. N)`）：缺省 / `<= 0` 一律 0（不重复），
 * 越界夹到行数。渲染、量测、分页、导出四处都走这一个函数 —— 解析与 normalizeTable
 * 虽然已经夹过一次，但那管不到手搓模型与旧数据，读侧再夹一次才不会冒出第二套口径
 * （两套口径的后果是「预览重复 1 行、docx 重复 2 行」这种两边对不上的分家）。
 */
export function headerRowCount(table: TableBlock): number {
  const raw = table.headerRows
  if (typeof raw !== 'number' || !Number.isFinite(raw)) return 0
  return Math.max(0, Math.min(Math.floor(raw), table.rows.length))
}

/**
 * 单元格在编辑层里的「伪块」id，**恒指格内的第 0 段**。
 *
 * 单元格也挂 data-block-id，于是 edit/dom.ts 的坐标换算（fragmentOf 向上取最近的
 * data-block-id）不用改一行就能把光标落进格子里 —— 代价只是编辑层按 id 找容器。
 * 第 N（N > 0）段的 id 见 cellParagraphId()。
 */
export function cellId(tableId: string, r: number, c: number): string {
  return `${tableId}.r${r}c${c}`
}

/**
 * 格内第 para 段的伪块 id。para 为 0 时就是 cellId(tableId, r, c) 本身
 * —— 既有的格子 id 与断言一字不改，第 2 段起才带 `.pN` 后缀。
 */
export function cellParagraphId(tableId: string, r: number, c: number, para: number): string {
  return para === 0 ? cellId(tableId, r, c) : `${cellId(tableId, r, c)}.p${para}`
}

const CELL_ID_RE = /^(.*)\.r(\d+)c(\d+)(?:\.p(\d+))?$/

/** cellId() / cellParagraphId() 的反向解析；不是单元格 id 时返回 null。`para` 缺省为 0 */
export function parseCellId(
  id: string,
): { tableId: string; row: number; col: number; para: number } | null {
  const m = CELL_ID_RE.exec(id)
  const tableId = m?.[1]
  if (m === null || !tableId) return null
  return {
    tableId,
    row: Number(m[2]),
    col: Number(m[3]),
    para: m[4] === undefined ? 0 : Number(m[4]),
  }
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

/**
 * 编辑器的两个开关（写进 md 文档首行的 `::editor`）。
 * 默认值**不落字段**（与围栏的 `minLines=1` / `cantSplit=yes` 同一个约定）：
 * `trackChanges` 缺省 false、`nav` 缺省 true。
 */
export interface EditorSettings {
  /** 修订模式 */
  trackChanges?: boolean
  /** 导航窗格 */
  nav?: boolean
}

/** 解析成实际值（缺省补齐）。调用方要的是「开关到底是开还是关」，不是字段有没有 */
export interface EditorFlags {
  trackChanges: boolean
  nav: boolean
}

export function resolveEditorFlags(editor?: EditorSettings): EditorFlags {
  return { trackChanges: editor?.trackChanges === true, nav: editor?.nav !== false }
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
  /** 编辑器开关（文档级，不是某一节），来源与去处都是 md 首行的 `::editor` */
  editor?: EditorSettings
}

export const emptyDoc = (): DocModel => ({ blocks: [], comments: [] })

/**
 * 载入时的规范化：**零块文档补一个空白正文段落**。
 *
 * 编辑层是靠页面上带 `data-block-id` 的片段把 DOM 读回模型的；零块 ⇒ 页面一个片段都没有
 * ⇒ 读回的循环体一次都不跑。而 `.wtp-content` 自己就是 contenteditable，浏览器把字插在它
 * 下面（不在任何片段里），于是字只留在 DOM 上：看得见、保存与导出却是空的，还不报错。
 * 补一个空段落就有了落点，`getModel()` / `toMd()` 才跟得上版面上打的字。
 *
 * 只碰零块这一种输入：非零块一律**原样返回**（一个块都不多不少，comments / editor 等字段
 * 一个都不丢）。不就地改输入 —— `emptyDoc()` 的返回值可能是调用方持有的引用。
 */
export function ensureBodyBlock(doc: DocModel): DocModel {
  if (doc.blocks.length > 0) return doc
  return {
    ...doc,
    blocks: [{ t: 'textBlock', id: nextBlockId(), kind: 'body', inlines: [] }],
  }
}

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
 * 全篇所有可编辑容器的 inlines（段落 + 表格每个格子的**每一段**），按文档顺序。
 *
 * 修订标记、批注锚点都可能落在格子里 —— 而且是格子里的任何一段，所以全篇扫描
 *（找修订 id 最大值、清批注锚点……这类事）必须走这一条：只遍历 blocks 里能看到的
 * textBlock 会漏，只取每格第 0 段同样会漏。
 * 模型坏掉（格子没有 paragraphs 字段）时跳过这一格，不抛。
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
      for (const cell of row.cells) {
        for (const para of cell.paragraphs ?? []) out.push(para)
      }
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

/**
 * **编辑用**的切片：与 sliceInlines 同一个坐标系，边界规则不同。
 *
 * 批注锚点：sliceInlines 是给分页用的，无条件保留所有锚点，好让跨页的每个片段都能把高亮画全。
 * 编辑读回不能那样 —— 把 `[from,to)` 的片段替换掉以后，`[0,from)` 与 `[to,len)` 里残留的锚点
 * 会变成重复的批注标记。所以这里按位置过滤：起点落在 `[from,to)` 内才留，终点落在 `(from,to]` 内才留。
 *
 * 软换行（零宽）：这里落在 `[from,to)` 就保住 —— 编辑读回没有行盒可丢，保住它比丢掉好
 * （切分段落时不会把换行弄没）；sliceInlines 反过来要在片首/片尾丢掉边界上那一枚。
 *
 * **它住在这里而不是 `edit/model.ts`**：编辑层（切段、替换）与表格结构层（格内分段，
 * 见 `edit/table.ts`）都要用它，放在 model.ts 会让那两个模块互相 import 成环。
 * 这里是模型的纯函数层，谁都能引。
 */
export function sliceStrict(inlines: readonly Inline[], from: number, to: number): Inline[] {
  const out: Inline[] = []
  let cursor = 0
  for (const inline of inlines) {
    if (inline.t === 'commentStart') {
      if (cursor >= from && cursor < to) out.push(inline)
      continue
    }
    if (inline.t === 'commentEnd') {
      if (cursor > from && cursor <= to) out.push(inline)
      continue
    }
    if (inline.t === 'break') {
      if (cursor >= from && cursor < to) out.push(inline)
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
