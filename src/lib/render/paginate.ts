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

/** 换页的原因：分页符只换页，分节符换页且新开一节（页码可重排） */
export type BreakKind = 'page' | 'section'

export interface MeasuredBreak {
  t: 'break'
  /** 模型里那个分页符/分节符块的 id，编辑器要按它删除 */
  blockId: string
  kind: BreakKind
  /** 仅分节符有意义：本节页码是否从 1 重新开始 */
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

/** 一枚换页标记（页底要画出来的那条） */
export interface PageBreakMark {
  kind: BreakKind
  /** 模型里那个分页符/分节符块的 id，编辑器要按它删除 */
  blockId: string
}

export interface PageLayout {
  sectionIndex: number
  /** 本节内页码，从 1 开始；是否重启用分节符控制 */
  pageNumber: number
  fragments: PageFragment[]
  /**
   * 这一页底部的换页标记，按出现顺序。
   *
   * 是**列表**而不是单值：「分页符 + 分节符」连在一起时，两枚标记都落在这张页
   * 的底部（Word 也把两条标记都画出来），只是其中一枚的换页被吸收而已 ——
   * 用单值就只能显示一个。文末为空的那一节也会推一张没有片段的空白页（Word 实测）。
   */
  breaks: PageBreakMark[]
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
      pages.push({ sectionIndex, pageNumber, fragments, breaks: [] })
      fragments = []
      pageNumber += 1
    }
    cursor = 0
  }

  /** 换页标记落下时，推进节号与页码。分节符新开一节，页码是否重排由它自己说了算。 */
  const stepNumbering = (item: MeasuredBreak): void => {
    if (item.kind === 'section') {
      sectionIndex += 1
      pageNumber = item.restartNumbering ? 1 : pageNumber + 1
    } else {
      pageNumber += 1
    }
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

  /**
   * 还没能找到落点的标记。只有一种情况会积在这里：整篇文档一开头就是换页标记 ——
   * 那时还没有任何一页可挂，先记着，等第一页出现时一起画上去。
   */
  const pending: PageBreakMark[] = []

  for (const item of items) {
    if (item.t === 'break') {
      const mark: PageBreakMark = { kind: item.kind, blockId: item.blockId }
      if (fragments.length > 0) {
        // 当前页有内容：换页成立，标记画在这一页底部
        pages.push({
          sectionIndex,
          pageNumber,
          fragments,
          breaks: [...pending, mark],
        })
        pending.length = 0
        fragments = []
      } else {
        // 当前页是空的：换页被吸收（Word 实测如此：分页符 + 分节符连在一起只推进一页），
        // 但标记不能丢 —— 挂到刚结束的那一页底部，两枚就都能看见了
        const last = pages[pages.length - 1]
        if (last) last.breaks.push(mark)
        else pending.push(mark)
      }
      stepNumbering(item)
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
    pages.push({ sectionIndex, pageNumber, fragments, breaks: [...pending] })
  } else if (
    pages.length === 0 ||
    (pages[pages.length - 1]?.sectionIndex ?? -1) !== sectionIndex
  ) {
    // 这一节一页都还没有：文末是分节符，或者整篇为空。
    // Word 会给这样的节留一张空白页（实测：内容 + 文末分节符 → Word 报 2 页 2 节），
    // 不补的话预览会比 Word 少一页，而「预览页数 = Word 页数」正是这个组件的前提。
    pages.push({ sectionIndex, pageNumber, fragments: [], breaks: [...pending] })
  }

  return pages
}
