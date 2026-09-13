/**
 * 编辑操作的模型层（纯函数，不碰 DOM，可在 node 里单测）。
 *
 * 编辑采用「DOM 是手感的真相、模型是导出的真相」这一分工：
 * 浏览器负责把字敲进去（原生编辑、输入法、原生选区都留着），
 * 每次 input 再把 DOM 读回模型；结构性操作（回车、退格合并、工具栏格式化、
 * 批注）反过来直接改模型，再重排渲染。两者都收敛到这一层，
 * 因此 docx 导出的永远是模型，不会出现「界面改了、导出没改」。
 *
 * 坐标系：这里的偏移一律是**模型文字坐标**（不含标题自动编号前缀）。
 * 预览 DOM 用的是「显示坐标」（含前缀），换算只在 dom.ts 里做。
 */

import type { Align, BlockKind } from '../spec'
import type {
  Block,
  CellVerticalAlign,
  CommentDef,
  DocModel,
  Inline,
  InlineHolder,
  PageOrientation,
  RevMark,
  TableRowRole,
  TextBlock,
} from '../types'
import { allInlineHolders, inlinesText, nextBlockId, parseCellId, plainText } from '../types'
import { findCell, setCellKind } from './table'
import { insertSectionBreakAfter, removeSectionBreak } from './section'

export interface BlockPoint {
  blockId: string
  /** 模型文字坐标下的字符偏移 */
  offset: number
}

/**
 * 光标落在表格格子里时的上下文，供调用方的「上下文工具条」回显与禁用按钮。
 *
 * 调用方（App）手里只有这一次 `selection-change` 事件，没有响应式的模型，
 * 所以行/列下标、行数、role 这些都得由组件现算后一并带出来。
 */
export interface TableSelectionContext {
  tableId: string
  /** 光标所在行在 rows 数组里的下标 */
  row: number
  /** 列下标；unit/note 行天然只有一格，恒为 0 */
  col: number
  role: TableRowRole
  /** 总行数（含 unit/note） */
  rows: number
  columns: number
  /** body 行的行数 */
  bodyRows: number
  minLines: 1 | 2
  hasUnit: boolean
  hasNote: boolean
  /** 已按角色默认值解析过的实际水平对齐；按钮高亮要按它，不按有没有覆盖 */
  alignH: Align
  /** 已解析默认值（缺省 top）的实际垂直对齐 */
  alignV: CellVerticalAlign
}

/** 工具栏要的选区信息（模型文字坐标，不含自动编号前缀） */
export interface EditorSelection {
  blockId: string
  kind: BlockKind
  from: number
  to: number
  /** 是否只是一个插入符（没有选中文字） */
  collapsed: boolean
  /** 选区内文字是否全加粗 */
  bold: boolean
  /** 选区内文字是否同色；混色或无色时为 undefined */
  color?: string
  /** 落点在表格格子里时给出表格上下文；不在格子里则没有这个字段 */
  table?: TableSelectionContext
  /** 光标所在节的上下文；供「节」上下文工具条回显与置灰 */
  section?: SectionSelectionContext
}

/**
 * 光标所在节的上下文。三个开关给的都是**已解析值**（缺省补齐、首节的 linkPrevious
 * 已强制 false），App 侧据此画 radio 的选中态与禁用态 —— App 手里没有响应式的模型，
 * 只能吃这一次 emit。
 */
export interface SectionSelectionContext {
  /** 节号（0 = 首节） */
  index: number
  /** 总节数 */
  total: number
  isFirst: boolean
  orientation: PageOrientation
  pageNumbers: boolean
  linkPrevious: boolean
  restartAtOne: boolean
}

/**
 * 能承载一段可编辑文字的东西 —— 唯一要求就是有一串 inlines。
 *
 * 「块」不止 textBlock：表格的每个格子也是一段可编辑文字（单段落），
 * 编辑层把它当成一个「伪块」（id 用 cellId(tableId, r, c)）。所以本文件里的
 * 编辑操作统一吃**容器**，段落与格子走同一条路，不必各写一套。
 */
export type InlineContainer = InlineHolder

export function findBlock(doc: DocModel, id: string): TextBlock | undefined {
  for (const block of doc.blocks) {
    if (block.t === 'textBlock' && block.id === id) return block
  }
  return undefined
}

/**
 * 按 id 找一个可编辑容器：段落按 id 命中；`tableId.rNcM` 按 cellId 命中最深的那个格子。
 *
 * 找不到（id 过期、格号越界、或者 id 其实是一张表的 id —— 表格本身不可编辑）返回 undefined，
 * 调用方据此安全跳过。
 */
export function findContainer(doc: DocModel, id: string): InlineContainer | undefined {
  const cell = parseCellId(id)
  for (const block of doc.blocks) {
    if (cell) {
      if (block.t !== 'table' || block.id !== cell.tableId) continue
      return block.rows[cell.row]?.cells[cell.col]
    }
    if (block.t === 'textBlock' && block.id === id) return block
  }
  return undefined
}

export function findBlockIndex(doc: DocModel, id: string): number {
  return doc.blocks.findIndex((block) => block.t === 'textBlock' && block.id === id)
}

/** 容器里的文字长度。段落与格子的长度都按这个算（软换行零宽、不占字符位） */
export function containerLength(container: InlineContainer): number {
  return inlinesText(container.inlines).length
}

export function blockLength(block: TextBlock): number {
  return plainText(block).length
}

/**
 * 按字符区间切 inline 片段，批注锚点跟着它夹住的那段文字走。
 *
 * 与 types.ts 的 sliceInlines 的差别：那个是给分页用的 —— 它无条件保留所有
 * 批注锚点，好让跨页的每个片段都能把高亮画全。编辑读回时不能那样：
 * 把 [from,to) 的片段替换掉以后，[0,from) 与 [to,len) 里残留的锚点会变成
 * 重复的批注标记。所以这里按锚点的位置过滤：起点落在 [from,to) 内才留，
 * 终点落在 (from,to] 内才留。
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
      // 零宽，只能按位置归属：落在 [from,to) 就跟着这一片走。
      // 与分页用的 sliceInlines 不同 —— 那里要在片首丢掉边界上那一枚（避免多出一个没记账的行盒），
      // 编辑读回没有行盒可丢，保住它比丢掉好（切分段落时不会把换行弄没）。
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

/**
 * 用新片段替换容器内 [from,to) 的文字（其余部分原样保留）。
 *
 * 容器的类型是「有 inlines 的东西」—— 段落与表格格子共用这一条路
 * （格子的单段落语义与段落一样，只有 id 的坐标系不同）。
 *
 * 批注锚点不能像 sliceStrict 那样按区间一刀切：夹住被替换文字的锚点，起点必须留在
 * 替换内容之前、终点必须留在之后。否则会剩下一个没有终点的 commentStart，
 * 渲染时它会把这一段余下的文字全吞进高亮里，导出 docx 也会写出没闭合的批注范围。
 * 所以中间的锚点按「夹到边界」处理：起点并入前半、终点并入后半，新内容自然被包住。
 * 两条规则（前半 cursor < a、后半 cursor >= b）互斥，插入（from == to）时也不会重复。
 *
 * 软换行（零宽）走同一条「前半 cursor < a、后半 cursor >= b」规则：
 * 替换区间内部的换行跟着区间一起被替换掉，边界上的不会重复或丢失。
 */
export function replaceRange(
  container: InlineContainer,
  from: number,
  to: number,
  inlines: readonly Inline[],
): void {
  const len = containerLength(container)
  const a = Math.max(0, Math.min(from, len))
  const b = Math.max(a, Math.min(to, len))
  const out: Inline[] = []

  let cursor = 0
  for (const inline of container.inlines) {
    if (inline.t !== 'text') {
      if (cursor < a || (cursor < b && inline.t === 'commentStart')) out.push(inline)
      continue
    }
    const start = cursor
    cursor += inline.text.length
    if (start >= a) continue
    const text = inline.text.slice(0, Math.min(a, cursor) - start)
    if (text !== '') out.push({ ...inline, text })
  }

  out.push(...inlines)

  cursor = 0
  for (const inline of container.inlines) {
    if (inline.t !== 'text') {
      if (cursor >= b || (cursor > a && cursor < b && inline.t === 'commentEnd')) {
        out.push(inline)
      }
      continue
    }
    const start = cursor
    cursor += inline.text.length
    if (cursor <= b) continue
    out.push({ ...inline, text: inline.text.slice(Math.max(b, start) - start) })
  }

  container.inlines = out
}

/** 在 offset 处插入纯文本（可选地带上修订标记） */
export function insertText(
  doc: DocModel,
  containerId: string,
  offset: number,
  text: string,
  rev?: RevMark,
): void {
  const container = findContainer(doc, containerId)
  if (!container || text === '') return
  const piece: Inline = { t: 'text', text, ...(rev ? { rev } : {}) }
  replaceRange(container, offset, offset, [piece])
}

/**
 * 删除 [from,to)。`rev` 非空时不真删，而是把删掉的文字标成 w:del 留在原处
 * ——这正是 Word 的「修订模式」在界面上看到的效果。
 */
export function deleteRange(
  doc: DocModel,
  containerId: string,
  from: number,
  to: number,
  rev?: RevMark,
): void {
  const container = findContainer(doc, containerId)
  if (!container || to <= from) return
  if (!rev) {
    replaceRange(container, from, to, [])
    return
  }
  // 保留原文字，只把区间内的 text inline 盖上删除标记（原来的 ins 标记要让位）
  const kept = sliceStrict(container.inlines, from, to).map((inline): Inline => {
    if (inline.t !== 'text') return inline
    return { ...inline, rev: { ...rev } }
  })
  replaceRange(container, from, to, kept)
}

/** 把块在 offset 处切成两块，后半段成为 tailKind 的新块；返回新块 id */
export function splitBlock(
  doc: DocModel,
  blockId: string,
  offset: number,
  tailKind: BlockKind,
): string {
  const index = findBlockIndex(doc, blockId)
  const block = doc.blocks[index]
  if (index < 0 || !block || block.t !== 'textBlock') return ''
  const len = blockLength(block)
  const at = Math.max(0, Math.min(offset, len))
  const tail = sliceStrict(block.inlines, at, len)
  block.inlines = sliceStrict(block.inlines, 0, at)
  const created: TextBlock = { t: 'textBlock', id: nextBlockId(), kind: tailKind, inlines: tail }
  doc.blocks.splice(index + 1, 0, created)
  return created.id
}

/** 把当前块并入前一块；返回合并后的落点（前一块的末尾）。前一块是分节符时不合并。 */
export function mergeIntoPrevious(doc: DocModel, blockId: string): BlockPoint | null {
  const index = findBlockIndex(doc, blockId)
  const block = doc.blocks[index]
  const prev = doc.blocks[index - 1]
  if (index <= 0 || !block || !prev || block.t !== 'textBlock' || prev.t !== 'textBlock') {
    return null
  }
  const at = blockLength(prev)
  prev.inlines = [...prev.inlines, ...block.inlines]
  doc.blocks.splice(index, 1)
  return { blockId: prev.id, offset: at }
}

/** 删除一个块（退格删空段、或块被并入后清理用） */
export function removeBlock(doc: DocModel, blockId: string): void {
  const index = findBlockIndex(doc, blockId)
  if (index >= 0) doc.blocks.splice(index, 1)
}

export function setBlockKind(doc: DocModel, blockId: string, kind: BlockKind): void {
  const block = findBlock(doc, blockId)
  if (block) block.kind = kind
}

/**
 * 给「任意可编辑容器」设样式：cellId 命中格子就设格子的 kind，否则设段落。
 *
 * 与只认 textBlock 的 setBlockKind 并存：那个的既有行为不许动
 * （单测与接口都在用），跨段落 + 跨格的工具栏操作走这一条。
 */
export function setContainerKind(doc: DocModel, containerId: string, kind: BlockKind): void {
  const cell = findCell(doc, containerId)
  if (cell) {
    setCellKind(cell, kind)
    return
  }
  setBlockKind(doc, containerId, kind)
}

/** 判断 [from,to) 内的文字是否全部加粗（空区间返回 false） */
export function rangeIsBold(container: InlineContainer, from: number, to: number): boolean {
  if (to <= from) return false
  let seen = false
  let cursor = 0
  for (const inline of container.inlines) {
    if (inline.t !== 'text') continue
    const start = cursor
    const end = cursor + inline.text.length
    cursor = end
    if (end <= from || start >= to) continue
    seen = true
    if (!inline.bold) return false
  }
  return seen
}

/** 判断 [from,to) 内的文字是否全部带下划线（空区间返回 false） */
export function rangeIsUnderline(container: InlineContainer, from: number, to: number): boolean {
  if (to <= from) return false
  let seen = false
  let cursor = 0
  for (const inline of container.inlines) {
    if (inline.t !== 'text') continue
    const start = cursor
    const end = cursor + inline.text.length
    cursor = end
    if (end <= from || start >= to) continue
    seen = true
    if (!inline.underline) return false
  }
  return seen
}

/** 判断 [from,to) 内的文字是否全是同一个颜色；不一致返回 undefined */
export function rangeColor(
  container: InlineContainer,
  from: number,
  to: number,
): string | undefined {
  if (to <= from) return undefined
  let found: string | undefined
  let seen = false
  let cursor = 0
  for (const inline of container.inlines) {
    if (inline.t !== 'text') continue
    const start = cursor
    const end = cursor + inline.text.length
    cursor = end
    if (end <= from || start >= to) continue
    if (!seen) {
      seen = true
      found = inline.color
    } else if (found !== inline.color) {
      return undefined
    }
  }
  return seen ? found : undefined
}

/**
 * 给 [from,to) 套格式。颜色传 null 表示「恢复默认色」（清掉 color）。
 * 加粗传 true/false，未传的字段保持原样。
 * 下划线同理：true 加上，false 与 null 都是清掉。
 */
export function applyFormat(
  doc: DocModel,
  containerId: string,
  from: number,
  to: number,
  patch: { bold?: boolean; underline?: boolean | null; color?: string | null },
): void {
  const container = findContainer(doc, containerId)
  if (!container || to <= from) return
  replaceRange(
    container,
    from,
    to,
    sliceStrict(container.inlines, from, to).map((inline): Inline => {
      if (inline.t !== 'text') return inline
      const next: Inline = { ...inline }
      if (patch.bold !== undefined) {
        if (patch.bold) next.bold = true
        else delete next.bold
      }
      if (patch.underline !== undefined) {
        if (patch.underline) next.underline = true
        else delete next.underline
      }
      if (patch.color !== undefined) {
        if (patch.color) next.color = patch.color
        else delete next.color
      }
      return next
    }),
  )
}

/** 给 [from,to) 加一条批注；返回新批注 id */
export function addComment(
  doc: DocModel,
  containerId: string,
  from: number,
  to: number,
  text: string,
  author: string,
  date: string,
): number {
  const container = findContainer(doc, containerId)
  if (!container || to <= from) return -1
  const id = doc.comments.reduce((max, c) => Math.max(max, c.id), -1) + 1
  const def: CommentDef = { id, author, date, text }
  doc.comments.push(def)
  // 用 replaceRange 落锚点：它会顺带把与新区间相交的旧锚点配对好（见该函数的说明）
  replaceRange(container, from, to, [
    { t: 'commentStart', commentId: id },
    ...sliceStrict(container.inlines, from, to),
    { t: 'commentEnd', commentId: id },
  ])
  return id
}

/** 回复某条批注（Word 里表现为同一线程下的追加） */
export function replyComment(
  doc: DocModel,
  parentId: number,
  text: string,
  author: string,
  date: string,
): number {
  const id = doc.comments.reduce((max, c) => Math.max(max, c.id), -1) + 1
  doc.comments.push({ id, author, date, text, parentId })
  return id
}

/** 改写某条批注的内容。作者与时间保持原样（改的是内容，不是“谁在什么时候说的”）。 */
export function updateComment(doc: DocModel, commentId: number, text: string): boolean {
  const target = doc.comments.find((c) => c.id === commentId)
  if (!target) return false
  target.text = text
  return true
}

/**
 * 在 blockId 之后插入一个分页符或分节符，返回新块的 id。
 * blockId 找不到（或没给）时追加到文末。
 *
 * 分节符走 edit/section.ts 的封装 —— 它还要同步 `doc.sections`，别在这里另写一份。
 */
export function insertBreakAfter(
  doc: DocModel,
  blockId: string | undefined,
  kind: 'page' | 'section',
): string {
  if (kind === 'section') return insertSectionBreakAfter(doc, blockId)
  const found = blockId === undefined ? -1 : doc.blocks.findIndex((b) => b.id === blockId)
  const at = found < 0 ? doc.blocks.length : found + 1
  const block: Block = { t: 'pageBreak', id: nextBlockId('pg') }
  doc.blocks.splice(at, 0, block)
  return block.id
}

/**
 * 删掉一个换页标记。**只认分页符 / 分节符** —— 段落不归它管（那是 removeBlock / 退格合并），
 * 表格更不归它管（删整表走 edit/table.ts 的 removeTable）。
 *
 * 早先这里只拒 textBlock，于是拿一张表的 id 调进来会把整张表静默删掉 —— 是个陷阱，
 * 现已收紧：其余块一律返回 false 且一个字节都不改。
 * 删分节符还要同步摘掉 `doc.sections` 里对应的那一项（转交 edit/section.ts）。
 */
export function removeBreak(doc: DocModel, blockId: string): boolean {
  const index = doc.blocks.findIndex((b) => b.id === blockId)
  const block = doc.blocks[index]
  if (index < 0 || !block) return false
  if (block.t === 'pageBreak') {
    doc.blocks.splice(index, 1)
    return true
  }
  if (block.t === 'sectionBreak') return removeSectionBreak(doc, blockId)
  return false
}

/** 删除整条批注及其锚点（父批注连同回复一起删） */
export function removeComment(doc: DocModel, commentId: number): void {
  const doomed = new Set<number>([commentId])
  let grew = true
  while (grew) {
    grew = false
    for (const c of doc.comments) {
      if (c.parentId !== undefined && doomed.has(c.parentId) && !doomed.has(c.id)) {
        doomed.add(c.id)
        grew = true
      }
    }
  }
  doc.comments = doc.comments.filter((c) => !doomed.has(c.id))
  for (const holder of allInlineHolders(doc)) {
    holder.inlines = holder.inlines.filter(
      (inline) =>
        (inline.t !== 'commentStart' && inline.t !== 'commentEnd') || !doomed.has(inline.commentId),
    )
  }
}

/** 深拷贝模型（撤销栈、渲染快照都用它，避免共享引用被后续编辑改到） */
export function cloneDoc(doc: DocModel): DocModel {
  return {
    blocks: doc.blocks.map((block): Block => {
      // 表格块要逐层新建：浅拷贝会让副本与原稿共享 rows/cells/inlines，
      // 撤销栈与渲染快照复原时会被后续编辑连带改到。
      if (block.t === 'table') {
        return {
          t: 'table',
          id: block.id,
          columns: block.columns,
          minLines: block.minLines,
          cantSplit: block.cantSplit,
          rows: block.rows.map((row) => ({
            role: row.role,
            cells: row.cells.map((cell) => ({
              inlines: cell.inlines.map((inline): Inline => ({ ...inline })),
              // kind / align 也要逐格拷：少拷一个字段，撤销与渲染快照就会「回到默认样式 / 默认对齐」
              ...(cell.kind !== undefined ? { kind: cell.kind } : {}),
              ...(cell.align !== undefined ? { align: { ...cell.align } } : {}),
            })),
          })),
        }
      }
      if (block.t !== 'textBlock') return { ...block }
      return {
        t: 'textBlock',
        id: block.id,
        kind: block.kind,
        inlines: block.inlines.map((inline): Inline => ({ ...inline })),
      }
    }),
    comments: doc.comments.map((c) => ({ ...c })),
    // sections 是文档级数组，漏拷会让撤销 / 渲染快照与编辑中的模型共享引用，
    // 于是「切到别的节改方向」会连带改到快照，撤销也回不去
    ...(doc.sections ? { sections: doc.sections.map((s) => ({ ...s })) } : {}),
  }
}
