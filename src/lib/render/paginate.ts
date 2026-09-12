/**
 * 分页算法（纯函数，不碰 DOM，因此可以在 node 里直接单测）。
 *
 * 为什么自己写而不用 paged.js 之类：它们靠重建 DOM 实现分页，会摧毁正在输入的
 * 插入符；而且分节重编号在命名页里很难表达。这里的行高是实测出来的确定量，
 * 分页结果可以算出来而不是猜出来。
 *
 * 孤行控制（widow/orphan）：默认开启，页尾与页首各至少留 2 行 ——
 * 也就是少于 4 行的段落不会被拆开，与 Word 默认行为一致。
 */

import type { BlockKind } from '../spec'

export interface MeasuredBlock {
  t: 'block'
  blockId: string
  kind: BlockKind
  /** 显示文字长度（含自动编号前缀）。分页偏移都用这个坐标系。 */
  displayLength: number
  /** 实测行数 */
  rows: number
  /** 实测行高（px） */
  lineHeight: number
  /** 段前间距（px），由规格表换算而来 */
  spaceBefore: number
  /** 段后间距（px） */
  spaceAfter: number
  /** 每行起始处的字符偏移，长度等于 rows，首项恒为 0 */
  rowStarts: number[]
}

export interface MeasuredBreak {
  t: 'break'
  restartNumbering: boolean
}

export type MeasuredItem = MeasuredBlock | MeasuredBreak

export interface PageFragment {
  blockId: string
  kind: BlockKind
  /** 在「显示文字」坐标系里的起止字符偏移 */
  from: number
  to: number
  /** true 表示这是上一页同一段落后半截（页顶续排，不缩进、不叠段前距） */
  continuation: boolean
}

export interface PageLayout {
  sectionIndex: number
  /** 本节内页码，从 1 开始；是否重启用分节符控制 */
  pageNumber: number
  fragments: PageFragment[]
}

export interface PaginateOptions {
  /** 版心高度（px） */
  contentHeight: number
  /** 孤行控制，默认开启 */
  widowOrphan?: boolean
}

export function paginate(
  items: readonly MeasuredItem[],
  options: PaginateOptions,
): PageLayout[] {
  const contentHeight = options.contentHeight
  const widowOrphan = options.widowOrphan !== false

  const pages: PageLayout[] = []
  let sectionIndex = 0
  let pageNumber = 1
  let fragments: PageFragment[] = []
  let cursor = 0

  /** 收尾当前页。页码只在当前页确实有内容时才前进，避免留下空白页。 */
  const newPage = (): void => {
    if (fragments.length > 0) {
      pages.push({ sectionIndex, pageNumber, fragments })
      fragments = []
      pageNumber += 1
    }
    cursor = 0
  }

  const place = (item: MeasuredBlock, rowFrom: number, rowTo: number): void => {
    fragments.push({
      blockId: item.blockId,
      kind: item.kind,
      from: item.rowStarts[rowFrom] ?? 0,
      to: item.rowStarts[rowTo] ?? item.displayLength,
      continuation: rowFrom > 0,
    })
  }

  for (const item of items) {
    if (item.t === 'break') {
      // 分节：结束当前页，新节从新的一页开始
      if (fragments.length > 0) {
        pages.push({ sectionIndex, pageNumber, fragments })
        fragments = []
      }
      sectionIndex += 1
      pageNumber = item.restartNumbering ? 1 : pageNumber + 1
      cursor = 0
      continue
    }

    let rowFrom = 0
    while (rowFrom < item.rows) {
      // 页顶不叠加段前间距 —— 与 Word 的表现一致
      const before = cursor === 0 ? 0 : item.spaceBefore
      const rowsLeft = item.rows - rowFrom
      const blockHeight = rowsLeft * item.lineHeight
      const available = contentHeight - cursor - before

      if (blockHeight <= available) {
        place(item, rowFrom, item.rows)
        cursor += before + blockHeight + item.spaceAfter
        rowFrom = item.rows
        continue
      }

      const fit = item.lineHeight > 0 ? Math.floor(available / item.lineHeight) : 0
      let take = fit
      if (widowOrphan) {
        // 下一页至少留 2 行；本页不足 2 行就整段挪走（因此 3 行以下不会被拆）
        take = Math.min(fit, rowsLeft - 2)
        if (take < 2) take = 0
      }

      if (take <= 0) {
        if (cursor === 0) {
          // 版心连一行都放不下：兜底放一行，否则会死循环
          take = 1
        } else {
          newPage()
          continue
        }
      }

      place(item, rowFrom, rowFrom + take)
      rowFrom += take
      if (rowFrom < item.rows) newPage()
    }
  }

  if (fragments.length > 0) {
    pages.push({ sectionIndex, pageNumber, fragments })
  }
  if (pages.length === 0) {
    pages.push({ sectionIndex: 0, pageNumber: 1, fragments: [] })
  }

  return pages
}
