/**
 * 表格的结构操作（纯函数，可在 node 里单测）。
 *
 * 这里只改「一张给定的表」的内部形状，不碰 doc 的块顺序、不碰 DOM、不 import Vue ——
 * 「新锚点怎么按新下标重算」「撤销怎么记」都是组件层的事（WordPaper.vue 的 6 个表格操作）。
 *
 * 三条不变式，每改完 rows 都由 normalizeTable 收口：
 *   · columns = body 行的最大格数，且至少 1；
 *   · unit / note 行天然整行一格，只保留第 0 格（渲染与导出都只用 cells[0]）；
 *   · rows 的顺序就是显示顺序（unit → body… → note）。
 *
 * 注意 cellId 里嵌的是行/列**下标**（`tb1.r2c1`），所以任何增删都会让其后的格子 id
 * 整体位移 —— 这是 W4a 定的坐标方案的必然结果，不是缺陷；调用方必须按新下标重算锚点。
 */

import type { BlockKind } from '../spec'
import type { CellVerticalAlign, DocModel, TableBlock, TableCellModel, TableRowModel } from '../types'
import { parseCellId } from '../types'

/** 按 id 找表格块。id 既可以是表格自己的 id，也可以是格子的 cellId（tableId.rNcM） */
export function findTable(doc: DocModel, id: string): TableBlock | undefined {
  const cell = parseCellId(id)
  const tableId = cell ? cell.tableId : id
  for (const block of doc.blocks) {
    if (block.t === 'table' && block.id === tableId) return block
  }
  return undefined
}

/**
 * 删除整张表。id 既可以是表格 id 也可以是格子的 cellId（复用 findTable）。
 *
 * 专表专用：`edit/model.ts` 的 removeBreak 只管分页符 / 分节符，
 * 段落与表格都不归它管（这条收紧见 model.ts 的注释）。
 */
export function removeTable(doc: DocModel, id: string): boolean {
  const table = findTable(doc, id)
  if (!table) return false
  const index = doc.blocks.indexOf(table)
  if (index < 0) return false
  doc.blocks.splice(index, 1)
  return true
}

/** 按 cellId 找格子（纯函数，便于 emitSelection / 单测） */
export function findCell(doc: DocModel, cellIdValue: string): TableCellModel | undefined {
  const cell = parseCellId(cellIdValue)
  if (!cell) return undefined
  const table = findTable(doc, cell.tableId)
  return table?.rows[cell.row]?.cells[cell.col]
}

/** 设格子的样式；kind === 'listItem'（缺省语义）时删除字段，保持模型不存冗余值 */
export function setCellKind(cell: TableCellModel, kind: BlockKind): void {
  if (kind === 'listItem') delete cell.kind
  else cell.kind = kind
}

/** 设 / 清某一维的对齐覆盖；value === null 表示删掉这一维；两维都没了就删掉 align 字段 */
export function setCellAlign(
  cell: TableCellModel,
  part: 'h' | 'v',
  value: string | null,
): void {
  const align = { ...(cell.align ?? {}) }
  if (value === null) {
    if (part === 'h') delete align.h
    else delete align.v
  } else if (part === 'h') {
    align.h = value as NonNullable<typeof align.h>
  } else {
    align.v = value as CellVerticalAlign
  }
  if (align.h === undefined && align.v === undefined) delete cell.align
  else cell.align = align
}

/** 落点：目标格 + 落在格首还是格尾 */
export interface CellStep {
  row: number
  col: number
  at: 'start' | 'end'
}

/** 一行在跨格遍历里的「真实格」：body 行按 cells.length 数，unit/note 行整行算一格（col 恒为 0） */
function rowSlots(table: TableBlock): { row: number; col: number }[] {
  const out: { row: number; col: number }[] = []
  table.rows.forEach((row, r) => {
    if (row.role !== 'body') {
      out.push({ row: r, col: 0 })
      return
    }
    // 只数模型里真实存在的格：渲染时凑矩形补出来的幻影格没有模型容器
    // （findContainer 返回 undefined），打进去是死路 —— 见 PLAN 8 第 8 条②。
    for (let c = 0; c < row.cells.length; c += 1) out.push({ row: r, col: c })
  })
  return out
}

/** 行优先的下一格 / 上一格（遍历 rows 数组顺序：unit → body… → note）。没有下一格返回 null */
export function stepCell(
  table: TableBlock,
  row: number,
  col: number,
  dir: 'next' | 'prev',
): CellStep | null {
  const slots = rowSlots(table)
  const at = slots.findIndex((s) => s.row === row && s.col === col)
  if (at < 0) return null
  const next = slots[at + (dir === 'next' ? 1 : -1)]
  if (!next) return null
  return { row: next.row, col: next.col, at: dir === 'next' ? 'start' : 'end' }
}

/** 上 / 下的相邻格：同列；目标是 unit/note 行时列取 0；越界（含打到幻影格）返回 null */
export function verticalCell(
  table: TableBlock,
  row: number,
  col: number,
  dir: 'up' | 'down',
): CellStep | null {
  const r = row + (dir === 'up' ? -1 : 1)
  const target = table.rows[r]
  if (!target) return null
  if (target.role !== 'body') return { row: r, col: 0, at: 'start' }
  if (!target.cells[col]) return null
  return { row: r, col, at: 'start' }
}

function clampIndex(index: number, length: number): number {
  if (!Number.isFinite(index)) return 0
  return Math.min(Math.max(0, Math.trunc(index)), Math.max(0, length - 1))
}

function bodyRows(table: TableBlock): TableRowModel[] {
  return table.rows.filter((row) => row.role === 'body')
}

/** 所有 body 行在 rows 数组里的下标，按显示顺序 */
export function bodyRowIndexes(table: TableBlock): number[] {
  const out: number[] = []
  table.rows.forEach((row, index) => {
    if (row.role === 'body') out.push(index)
  })
  return out
}

/**
 * 把「想插在第几行」夹进 body 行的插入区间，使插进来的永远是 body 行。
 *
 * 不夹的话，光标停在 unit 行时点「上方插入行」会插到 unit 之前、停在 note 行时点「下方插入行」
 * 会插到 note 之后 —— 两者都会破坏 unit → body… → note 的显示顺序。
 * 返回的仍是 rows 数组下标的坐标系（可以直接交给 insertBodyRow）。
 */
export function bodyInsertIndex(table: TableBlock, at: number): number {
  const bodies = bodyRowIndexes(table)
  if (bodies.length > 0) {
    const first = bodies[0] as number
    const last = bodies[bodies.length - 1] as number
    return Math.min(Math.max(first, Math.trunc(at) || 0), last + 1)
  }
  // 一个 body 行都没有（md 里只写了 unit / note 的表）：插到 note 之前，没有 note 就插在末尾
  const note = table.rows.findIndex((row) => row.role === 'note')
  return note >= 0 ? note : table.rows.length
}

/** 新建一行 body：格数取「当时」的 columns（之后 normalizeTable 才会重算） */
function emptyBodyRow(columns: number): TableRowModel {
  return {
    role: 'body',
    cells: Array.from({ length: Math.max(1, columns) }, () => ({ inlines: [] })),
  }
}

/** 在 at 处（rows 数组下标，0..rows.length）插一行 body 行，返回新行的下标。格数 = 当时的 table.columns */
export function insertBodyRow(table: TableBlock, at: number): number {
  const index = clampIndex(at, table.rows.length + 1)
  table.rows.splice(index, 0, emptyBodyRow(table.columns))
  normalizeTable(table)
  return index
}

/** 删 at 处的 body 行。at 不是 body 行、或删完 body 行为 0 → 拒绝，返回 false */
export function removeBodyRow(table: TableBlock, at: number): boolean {
  const row = table.rows[at]
  if (!row || row.role !== 'body') return false
  if (bodyRows(table).length <= 1) return false
  table.rows.splice(at, 1)
  normalizeTable(table)
  return true
}

/** radio 开关：加/删 unit（表头行，插在最前）或 note（附注行，加在最后）。各至多一行；已存在时 on=true 为幂等空操作 */
export function setRoleRow(table: TableBlock, role: 'unit' | 'note', on: boolean): void {
  const has = table.rows.some((row) => row.role === role)
  if (on) {
    if (has) return
    const row: TableRowModel = { role, cells: [{ inlines: [] }] }
    if (role === 'unit') table.rows.unshift(row)
    else table.rows.push(row)
  } else {
    if (!has) return
    table.rows = table.rows.filter((row) => row.role !== role)
  }
  normalizeTable(table)
}

/** 在所有 body 行的 at 列处插一个空格；unit/note 行不动。at 按每行实际格数夹取（md 参差行不要求先拍平） */
export function insertColumn(table: TableBlock, at: number): void {
  for (const row of table.rows) {
    if (row.role !== 'body') continue
    row.cells.splice(clampIndex(at, row.cells.length + 1), 0, { inlines: [] })
  }
  normalizeTable(table)
}

/** 删所有 body 行的 at 列；columns 会掉到 0 → 拒绝并保持原样，返回 false；unit/note 行不动 */
export function removeColumn(table: TableBlock, at: number): boolean {
  const bodies = bodyRows(table)
  const index = clampIndex(at, table.columns)
  // 先算删完之后的「body 行最大格数」：会掉到 0 就整条拒绝，一格都不许动（调用方要能深比较出「原样」）
  const maxAfter = bodies.reduce(
    (max, row) => Math.max(max, index < row.cells.length ? row.cells.length - 1 : row.cells.length),
    0,
  )
  if (maxAfter <= 0) return false
  for (const row of bodies) {
    if (index < row.cells.length) row.cells.splice(index, 1)
  }
  normalizeTable(table)
  return true
}

/** columns = body 行最大格数（至少 1）；unit/note 行只保留第 0 格（渲染与导出都只用 cells[0]） */
export function normalizeTable(table: TableBlock): void {
  let columns = 1
  for (const row of table.rows) {
    if (row.role !== 'body') continue
    columns = Math.max(columns, row.cells.length)
  }
  table.columns = columns
  for (const row of table.rows) {
    if (row.role === 'body') continue
    row.cells = row.cells.slice(0, 1)
    if (row.cells.length === 0) row.cells.push({ inlines: [] })
  }
}

export function setMinLines(table: TableBlock, minLines: 1 | 2): void {
  table.minLines = minLines
}

/* -------------------------------------------------------------------------- */
/* 整格复选：矩形格区间、目标格列表、批量落笔（W7）                              */
/* -------------------------------------------------------------------------- */

/**
 * 一个矩形格区间（模型坐标：行列都是 rows / cells 的下标、闭区间）。
 *
 * 复选态是**组件的交互状态**，不进模型 —— 模型只在批量操作执行时被改。之所以存矩形而不是
 * 「选了哪些格」的平铺清单：同一张表跨页会渲成多个片段，高亮要按模型坐标下发，
 * 矩形是唯一同时能表达「一块」与「跨页」的形状（不连续块之间不做合并显示，见 PLAN 13.2）。
 */
export interface CellRect {
  r1: number
  c1: number
  r2: number
  c2: number
}

/** 一个格子的模型坐标 */
export interface CellRef {
  row: number
  col: number
}

/**
 * 指针落在的格子归一化成模型坐标：unit / note 行整行一格，列一律归到 0。
 *
 * 这两行在渲染与导出里都只有第 0 格（展开成 columnSpan = columns），不归一化的话
 * 「点在这两行的第 5 列」会指向一个不存在的格。
 */
export function normalizeCellCol(table: TableBlock, row: number, col: number): number {
  return table.rows[row]?.role === 'body' ? col : 0
}

/** 两点之间的矩形（行列各自取 min / max），与两个端点的先后顺序无关 */
export function cellRectBetween(a: CellRef, b: CellRef): CellRect {
  return {
    r1: Math.min(a.row, b.row),
    c1: Math.min(a.col, b.col),
    r2: Math.max(a.row, b.row),
    c2: Math.max(a.col, b.col),
  }
}

/**
 * 矩形里**真实存在**的格（按扫描顺序，未排序）。
 *
 * 越界的行列跳过（正常不会发生；撤销、结构操作之后可能）；body 行只数模型里真有的格 ——
 * 渲染时凑矩形补出来的幻影格没有容器，选中它没有意义（同 rowSlots 那条约定）。
 * unit / note 行只有第 0 格：矩形要覆盖第 0 列才算选中它（与 normalizeCellCol 同一条约定）。
 */
export function cellsInRect(table: TableBlock, rect: CellRect): CellRef[] {
  const r1 = Math.min(rect.r1, rect.r2)
  const r2 = Math.max(rect.r1, rect.r2)
  const c1 = Math.min(rect.c1, rect.c2)
  const c2 = Math.max(rect.c1, rect.c2)
  const out: CellRef[] = []
  for (let r = r1; r <= r2; r += 1) {
    const row = table.rows[r]
    if (!row) continue
    if (row.role !== 'body') {
      if (c1 <= 0 && c2 >= 0) out.push({ row: r, col: 0 })
      continue
    }
    for (let c = Math.max(0, c1); c <= c2 && c < row.cells.length; c += 1) {
      out.push({ row: r, col: c })
    }
  }
  return out
}

/** 行优先排序 + 去重。批量操作的遍历顺序必须稳定（一次批量 = 一步撤销） */
export function sortCells(cells: readonly CellRef[]): CellRef[] {
  const seen = new Set<string>()
  const out: CellRef[] = []
  for (const cell of cells) {
    const key = `${cell.row},${cell.col}`
    if (seen.has(key)) continue
    seen.add(key)
    out.push({ row: cell.row, col: cell.col })
  }
  return out.sort((a, b) => a.row - b.row || a.col - b.col)
}

/** 若干矩形块里的全部格（行优先、去重）—— 「选中了哪些格」由 blocks 现推 */
export function cellsInRects(table: TableBlock, rects: readonly CellRect[]): CellRef[] {
  return sortCells(rects.flatMap((rect) => cellsInRect(table, rect)))
}

/** 包含某格的块下标（Ctrl+点击已选中的格 = 去掉包含它的那一块）；没有就返回 -1 */
export function cellRectIndexOf(
  table: TableBlock,
  rects: readonly CellRect[],
  cell: CellRef,
): number {
  return rects.findIndex((rect) =>
    cellsInRect(table, rect).some((c) => c.row === cell.row && c.col === cell.col),
  )
}

/**
 * 批量对齐的下一步值：目标格**每一格**的该维实际生效值都等于 value → 清除覆盖（null），
 * 否则一律写 value。与单选「再点同一个值就是清除覆盖」是同一条规则，只是整批一起判。
 * 实际生效值由调用方算好传进来（覆盖 ?? 角色 / 样式默认 —— 渲染、导出、工具条回显都按这一条）。
 */
export function nextAlignValue(effective: readonly (string | null)[], value: string): string | null {
  return effective.length > 0 && effective.every((v) => v === value) ? null : value
}

/** 某格某一维对齐的**存储值**（没有覆盖就是 null）。批量改与「改了没有」的预判共用它 */
export function storedCellAlign(cell: TableCellModel, part: 'h' | 'v'): string | null {
  return (part === 'h' ? cell.align?.h : cell.align?.v) ?? null
}

/** 某格的**存储样式**（缺省语义 = listItem）。批量改与预判共用它 */
export function storedCellKind(cell: TableCellModel): BlockKind {
  return cell.kind ?? 'listItem'
}

/** 这一批格里「该维存储值与目标值不同」的格 —— 空数组 = 这一下是空转，调用方不该记撤销 */
export function cellsChangingAlign(
  table: TableBlock,
  cells: readonly CellRef[],
  part: 'h' | 'v',
  value: string | null,
): CellRef[] {
  return sortCells(cells).filter(({ row, col }) => {
    const cell = table.rows[row]?.cells[col]
    return cell !== undefined && storedCellAlign(cell, part) !== value
  })
}

/** 这一批格里「样式与目标样式不同」的格 —— 空数组 = 这一下是空转 */
export function cellsChangingKind(
  table: TableBlock,
  cells: readonly CellRef[],
  kind: BlockKind,
): CellRef[] {
  return sortCells(cells).filter(({ row, col }) => {
    const cell = table.rows[row]?.cells[col]
    return cell !== undefined && storedCellKind(cell) !== kind
  })
}

/** 批量落笔：把这一批格设成同一个对齐值（null = 清除该维覆盖） */
export function setCellsAlign(
  table: TableBlock,
  cells: readonly CellRef[],
  part: 'h' | 'v',
  value: string | null,
): void {
  for (const { row, col } of sortCells(cells)) {
    const cell = table.rows[row]?.cells[col]
    if (cell) setCellAlign(cell, part, value)
  }
}

/** 批量落笔：把这一批格设成同一个样式（设回缺省 listItem 就删字段） */
export function setCellsKind(table: TableBlock, cells: readonly CellRef[], kind: BlockKind): void {
  for (const { row, col } of sortCells(cells)) {
    const cell = table.rows[row]?.cells[col]
    if (cell) setCellKind(cell, kind)
  }
}
