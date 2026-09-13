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

import type { DocModel, TableBlock, TableRowModel } from '../types'
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
