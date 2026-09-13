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
import type { SectionRuntime } from '../section'

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

/** 换页的原因：分页符只换页，分节符换页且新开一节（页码是否重排由该节的设置决定） */
export type BreakKind = 'page' | 'section'

export interface MeasuredBreak {
  t: 'break'
  /** 模型里那个分页符/分节符块的 id，编辑器要按它删除 */
  blockId: string
  kind: BreakKind
}

/**
 * 表格的一行（已经量好高度）。
 *
 * 行是**原子**的：表格的断行只发生在行与行之间（Word 的 w:cantSplit 语义），
 * 行内不做切分，所以这里没有 rowStarts，只有实测行高。
 */
export interface MeasuredTableRow {
  t: 'tableRow'
  /** 表格块 id */
  blockId: string
  /** 行在 block.rows 里的下标 */
  row: number
  /** 实测行高（px） */
  height: number
}

export type MeasuredItem = MeasuredBlock | MeasuredBreak | MeasuredTableRow

/**
 * 一页里的一个片段。
 *
 * 分两种：段落片段（kind / from / to 有意义）与**表格片段** ——
 * 表格按行装箱，一段可以覆盖若干整行，所以它用 rowFrom / rowTo 表达，且
 * `kind` 为空：表格不是 BlockKind，硬塞一个（比如 listItem）会让「这是什么块」
 * 变成一句假话。判别方式就是 `rowFrom !== undefined`。
 */
export interface PageFragment {
  blockId: string
  /** 段落片段的样式类别；表格片段为空 */
  kind?: BlockKind
  /** 在「显示文字」坐标系里的起止字符偏移（表格片段置 0，字符区间对它无意义） */
  from: number
  to: number
  /** true 表示这是上一页同一段落后半截（页顶续排，不缩进、不叠段前距） */
  continuation: boolean
  /** 表格片段：本片覆盖的表格行区间 [rowFrom, rowTo)（半开）。段落片段没有这两个字段 */
  rowFrom?: number
  rowTo?: number
}

/** 是不是表格片段（`kind` 为空、带行区间） */
export function isTableFragment(frag: PageFragment): boolean {
  return frag.rowFrom !== undefined
}

/** 一枚换页标记（页底要画出来的那条） */
export interface PageBreakMark {
  kind: BreakKind
  /** 模型里那个分页符/分节符块的 id，编辑器要按它删除 */
  blockId: string
}

export interface PageLayout {
  sectionIndex: number
  /** 本节内页码，从 1 开始；是否重启用分节符 + 该节设置控制 */
  pageNumber: number
  /** 这一页最终显不显示页码（分页器只搬运解析好的结果，不解析「关联前节」） */
  showPageNumber: boolean
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

/** 一节缺省时的运行时参数：版心高 = 全局值、显示页码、不重排 */
function defaultRuntime(contentHeight: number): SectionRuntime {
  return { contentHeight, showPageNumber: true, restartAtOne: false }
}

export interface PaginateOptions {
  /** 版心高度（px），`sections` 缺项时的兜底 */
  contentHeight: number
  /** 逐节运行时参数（下标 = 节号）。缺项一律用 contentHeight / 显示页码 / 不重排兜底 */
  sections?: readonly SectionRuntime[]
  /** 孤行控制，默认开启 */
  widowOrphan?: boolean
}

export function paginate(
  items: readonly MeasuredItem[],
  options: PaginateOptions,
): PageLayout[] {
  const widowOrphan = options.widowOrphan !== false
  const runtimeOf = (index: number): SectionRuntime =>
    options.sections?.[index] ?? defaultRuntime(options.contentHeight)

  const pages: PageLayout[] = []
  let sectionIndex = 0
  let pageNumber = 1
  // 逐节版心高：横竖混排时每节不同，装箱只能按当前节算。换节时在 stepNumbering 里跟着换。
  let contentHeight = runtimeOf(0).contentHeight
  let fragments: PageFragment[] = []
  let cursor = 0

  /** 收尾当前页。页码只在当前页确实有内容时才前进，避免留下空白页。 */
  const newPage = (): void => {
    if (fragments.length > 0) {
      pages.push({
        sectionIndex,
        pageNumber,
        showPageNumber: runtimeOf(sectionIndex).showPageNumber,
        fragments,
        breaks: [],
      })
      fragments = []
      pageNumber += 1
    }
    cursor = 0
  }

  /** 换页标记落下时，推进节号与页码。分节符新开一节，页码是否重排由该节设置说了算。 */
  const stepNumbering = (item: MeasuredBreak): void => {
    if (item.kind === 'section') {
      sectionIndex += 1
      contentHeight = runtimeOf(sectionIndex).contentHeight
      pageNumber = runtimeOf(sectionIndex).restartAtOne ? 1 : pageNumber + 1
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
   * 放下一整行表格。
   *
   * **同页相邻的同表行必须合并成一个片段**：一页只出一张 `<table>`，
   * 每行一个片段就会渲出「一行一张表」，边框与列宽各算各的。
   * 跨页时必须断开（页间本来就不是同一个 `<table>`），所以只在「本页最后一个片段
   * 就是同一张表」时才往后延 rowTo，否则新起一个片段（continuation = 不是从第 0 行开始）。
   */
  const placeTableRow = (item: MeasuredTableRow): void => {
    const last = fragments[fragments.length - 1]
    if (last !== undefined && last.rowFrom !== undefined && last.blockId === item.blockId) {
      last.rowTo = item.row + 1
      return
    }
    fragments.push({
      blockId: item.blockId,
      from: 0,
      to: 0,
      continuation: item.row > 0,
      rowFrom: item.row,
      rowTo: item.row + 1,
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
          showPageNumber: runtimeOf(sectionIndex).showPageNumber,
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

    if (item.t === 'tableRow') {
      /*
       * 表格行的**独立分支**：行是原子的，不进孤行控制。
       *
       * 不能让它落进下面的 widow/orphan 算术：那里的 `take = Math.min(fit, rowsLeft - 2)`
       * 对单行项会算成 -1 → take=0 → 整项挪到下一页。「放不下就整行挪走」行为上恰好是对的，
       * 但那是副作用，不是承诺 —— 显式写出来才不会在别人动孤行控制时被带坏。
       */
      if (cursor > 0 && item.height > contentHeight - cursor) newPage()
      // cursor 已是 0 却仍放不下（一行比整页还高）：兜底放下去，否则会死循环
      placeTableRow(item)
      cursor += item.height
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
    pages.push({
      sectionIndex,
      pageNumber,
      showPageNumber: runtimeOf(sectionIndex).showPageNumber,
      fragments,
      breaks: [...pending],
    })
  } else if (
    pages.length === 0 ||
    (pages[pages.length - 1]?.sectionIndex ?? -1) !== sectionIndex
  ) {
    // 这一节一页都还没有：文末是分节符，或者整篇为空。
    // Word 会给这样的节留一张空白页（实测：内容 + 文末分节符 → Word 报 2 页 2 节），
    // 不补的话预览会比 Word 少一页，而「预览页数 = Word 页数」正是这个组件的前提。
    pages.push({
      sectionIndex,
      pageNumber,
      showPageNumber: runtimeOf(sectionIndex).showPageNumber,
      fragments: [],
      breaks: [...pending],
    })
  }

  return pages
}
