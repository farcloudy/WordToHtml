<script setup lang="ts">
/**
 * A4 分页预览 + 编辑层（P3）。
 *
 * 只读时：源码进，A4 分页版面出，可导出 docx。
 * 编辑时：DOM 是手感的真相，模型是导出的真相。
 *
 *   · 打字、输入法、选区、原生剪贴板都留给浏览器 —— 手感是原生的；
 *   · 每次 input 再把 DOM 读回模型（见 lib/edit/dom.ts、lib/edit/model.ts）；
 *   · 结构性操作（回车、退格合并、工具栏加粗／下划线／改色、批注）直接改模型，再重排渲染。
 *
 * 「正常输入不得触发重排」这条约束靠两点落实：
 *   1. 渲染只读 viewDoc / pages 这两个浅响应式快照，编辑中的模型（doc）不参与渲染，
 *      所以打字不会让 Vue 重新渲染，DOM 与插入符原封不动；
 *   2. 每次 input 仍然量测一遍（带增量缓存，只重量改过的那块），只有当分页结果
 *      真的变了才更新快照、重建 DOM，并按「块 id + 字符偏移」把插入符放回去。
 */

import { computed, nextTick, onBeforeUnmount, onMounted, ref, shallowRef, watch } from 'vue'

import { toBlob } from '../lib/docx/export'
import { formatAmount } from '../lib/edit/amount'
import {
  currentRange,
  displayPointOf,
  fragmentOf,
  offsetToPoint,
  placeCaret,
  placeCaretAfterBreak,
  placeRange,
  prefixLengthOf,
  readInlines,
  selectedRanges,
} from '../lib/edit/dom'
import type { DisplayPoint, DisplayRange } from '../lib/edit/dom'
import { buildOutline, outlineSignature } from '../lib/edit/outline'
import type { OutlineEntry } from '../lib/edit/outline'
import { findMatches, replaceMatches, validateQuery } from '../lib/edit/search'
import type { Match, SearchOptions, SearchScope } from '../lib/edit/search'
import {
  DEFAULT_SHORTCUTS,
  SHORTCUT_ACTIONS,
  matchShortcut,
  resolveShortcuts,
} from '../lib/edit/shortcuts'
import type { ShortcutAction, ShortcutOverrides, ShortcutTable } from '../lib/edit/shortcuts'
import {
  addComment,
  applyFormat,
  blockLength,
  canJoinWithNext,
  canMergeIntoPrevious,
  cloneDoc,
  containerLength,
  deleteRange,
  deleteSpan,
  findBlock,
  findContainer,
  hasRevisions,
  insertBreakAfter,
  insertText,
  joinWithNext,
  mergeIntoPrevious,
  rangeColor,
  rangeIsBold,
  rangeIsUnderline,
  removeBreak as removeBreakOp,
  removeComment as removeCommentOp,
  replyComment as replyCommentOp,
  replaceRange,
  resolveRevisions,
  setContainerKind as setContainerKindOp,
  sliceStrict,
  splitBlock,
  updateComment as updateCommentOp,
} from '../lib/edit/model'
import type {
  BlockPoint,
  CellSelectionSummary,
  EditorSelection,
  SectionSelectionContext,
} from '../lib/edit/model'
import {
  bodyInsertIndex,
  bodyRowIndexes,
  cellRectBetween,
  cellRectIndexOf,
  cellsChangingAlign,
  cellsChangingKind,
  cellsInRects,
  findCell,
  findTable,
  insertBodyRow,
  insertColumn,
  nextAlignValue,
  normalizeCellCol,
  removeBodyRow,
  removeColumn,
  removeTable as removeTableOp,
  setCellsAlign,
  setCellsKind,
  setMinLines,
  setRoleRow,
  sortCells,
  stepCell,
  storedCellKind,
  verticalCell,
} from '../lib/edit/table'
import type { CellRect, CellRef, CellStep } from '../lib/edit/table'
import { parseMd } from '../lib/md/parse'
import { computeNumbering } from '../lib/numbering'
import { resolveSections } from '../lib/section'
import type { ResolvedSection } from '../lib/section'
import {
  CELL_SELECTION_CLASS,
  KEEP_SELECTION_HIGHLIGHT,
  SEARCH_CURRENT_HIGHLIGHT,
  SEARCH_HIGHLIGHT,
  injectCss,
} from '../lib/render/css'
import { renderInlinesHtml, renderTableFragment } from '../lib/render/html'
import { clearMeasureCache, measureDocument } from '../lib/render/measure'
import type { MeasureCache } from '../lib/render/measure'
import { paginate } from '../lib/render/paginate'
import type { BreakKind, MeasuredItem, PageFragment, PageLayout } from '../lib/render/paginate'
import { contentBoxPx, resolveSpec } from '../lib/spec'
import type { Align, BlockKind, DeepPartial, Spec } from '../lib/spec'
import {
  allInlineHolders,
  cellId,
  commentScopes,
  defaultCellAlignH,
  nextBlockId,
  parseCellId,
  resolveEditorFlags,
  sliceInlines,
} from '../lib/types'
import type {
  CellVerticalAlign,
  DocModel,
  EditorFlags,
  EditorSettings,
  Inline,
  PageOrientation,
  RevMark,
  SectionSettings,
  TableBlock,
  TextBlock,
  TextInline,
} from '../lib/types'
import { sectionIndexOf, setSectionSetting } from '../lib/edit/section'

const props = withDefaults(
  defineProps<{
    /** 类 md 源码。与 model 二选一，model 优先 */
    source?: string
    /** 直接给文档模型（调用方已经解析过时用） */
    model?: DocModel
    /** 规格覆盖，例如 { page: { margin: { top: '30mm' } } } */
    spec?: DeepPartial<Spec>
    /** 修订与批注的作者名 */
    author?: string
    /** 打开编辑层（直接在 A4 版面上改） */
    editable?: boolean
    /** 修订模式：新增标 w:ins，删除不真删、标成 w:del 留在原处 */
    trackChanges?: boolean
    /**
     * 自定义快捷键表：只写要改的动作，留空即用默认表 ——
     * 默认值就是 `lib/edit/shortcuts.json` 那份文件（见 lib/edit/shortcuts.ts）。
     * 写成工厂：对象型默认值在 Vue 里本来就该按实例现造，免得把同一个对象发给所有实例。
     */
    shortcuts?: ShortcutOverrides
  }>(),
  {
    source: '',
    author: '管理员',
    editable: false,
    trackChanges: false,
    shortcuts: () => ({ ...DEFAULT_SHORTCUTS }),
  },
)

const emit = defineEmits<{
  paginated: [count: number]
  'selection-change': [selection: EditorSelection | null]
  /** 按了 ctrl+shift+E：修订模式由调用方持有，组件只报告「该翻转了」 */
  'toggle-track-changes': []
  /** 一句话提示（无效输入之类），组件不做提示 UI，交给调用方 */
  toast: [message: string]
  /** 按了 ctrl+F / ctrl+G：面板由调用方画，这里只报告该开哪一种 */
  'open-search': [mode: 'find' | 'replace']
  /** 匹配数与当前位置变了；error 非空时是非法正则 */
  'search-state': [state: { total: number; current: number; error: string | null }]
  /** 大纲变了（按指纹比对，同一份大纲只发一次），导航窗格用它 */
  'outline-change': [entries: OutlineEntry[]]
  /**
   * 模型里 `::editor` 解析出来的两个开关（已补齐默认值）。模型是权威 ——
   * 调用方据此设自己的「修订模式 / 导航」状态，载入时就不该用自己的默认值顶掉模型里的值。
   * 只在**模型载入 / 重建**之后发；调用方改动开关时走 setEditorFlags 回写（那条路不发事件，
   * 否则两边互相写会把用户刚改的值覆盖回去）。
   */
  'editor-flags': [flags: EditorFlags]
}>()

const resolved = computed<Spec>(() => resolveSpec(props.spec))

/**
 * 权威模型。用 shallowRef：它是「导出的真相」，但**不参与渲染** ——
 * 渲染只读 viewDoc 快照，否则每敲一个字都会重建 DOM、丢掉插入符。
 */
const doc = shallowRef<DocModel>(props.model ?? parseMd(props.source, { author: props.author }))

/** 渲染用的冻结快照，只在重排时整体换新 */
const viewDoc = shallowRef<DocModel>(doc.value)
const viewNumbering = shallowRef<Map<string, string>>(new Map())

/**
 * 实际生效的快捷键表（默认表 + 调用方的覆盖）。表是纯函数算出来的，
 * 所以做成 computed：调用方换表时重新解析一次，顺带把「未知动作名 / 冲突」的 warn 打出来。
 */
const shortcuts = computed<ShortcutTable>(() => resolveShortcuts(props.shortcuts))

const host = ref<HTMLElement | null>(null)
const root = ref<HTMLElement | null>(null)
const pages = shallowRef<PageLayout[]>([])
/** 最近一次分页实际用到的量测值，排错时用来和渲染结果对账 */
const measured = shallowRef<MeasuredItem[]>([])
/**
 * 逐节解析结果（下标 = 节号）。页面几何、页码显示、工具条都吃它 ——
 * 它是「模型 + 规格表」的纯函数产物，所以每次重排跟着 viewDoc 一起换新。
 */
const resolvedSections = shallowRef<ResolvedSection[]>([])
/** 侧栏里正在查看的批注；只影响高亮，不触发重排 */
const activeCommentId = ref<number | null>(null)
/** 侧栏里正在改写的那条批注与草稿（只在侧栏里用，不进模型） */
const editingId = ref<number | null>(null)
const editDraft = ref('')

const cache: MeasureCache = new Map()
let cssKey = ''
let token = 0
/** 输入法组字中：这段时间不读回模型、不重排，否则会把拼音打断 */
let composing = false
/** 最近一次落点（selectionchange 时记录），退格合并等结构性操作要用 */
let lastCaret: DisplayPoint | null = null
/** beforeinput 时记下的落点：那时 DOM 还没变，是「本次编辑之前」的位置 */
let preEditCaret: DisplayPoint | null = null
/** 最近一次「真的选中了文字」的区间。点工具栏/批注框会让焦点离开正文，
 *  那时再读实时选区就只剩输入框里的空选区了，加批注要靠这份记录。 */
let stickyRanges: DisplayRange[] = []
/** 组字结束的兜底读回（浏览器不补 input 事件时用） */
let compSyncTimer: ReturnType<typeof setTimeout> | null = null

/*
 * 表格整格复选（W7）。故意**不用响应式**，理由是它就是「不重排」本身：
 * 做成 ref 会让 v-html 的片段跟着重渲，格内 DOM 被重建一次（插入符、原生选区一起丢），
 * 而这条交互的全部要求恰恰是「刷选、Ctrl+点击、清空选中一律不许改 viewDoc / pages、
 * 不许触发重量测与重建 DOM」。所以状态放在模块级变量里，高亮由 paintCellSelection
 * 命令式地挂类名（只改 class，不动几何）。
 */
let cellSelection: CellSelection | null = null
/** 上一次整格操作的落点（Ctrl+点击追加时的锚格） */
let cellAnchor: CellRef | null = null
/** 正在按下的这一次指针手势；没有越出起点格之前它只是普通的选文字 */
let cellDrag: CellDrag | null = null

/* 查找会话。故意不用响应式：面板由调用方持有，组件只需要在数字/指纹变化时 emit，
   而这些状态每敲一个字都可能变，做成 ref 反而会让 Vue 白白重渲染。 */
let searchOn = false
let searchQuery = ''
let searchRegex = false
let searchScope: SearchScope[] | undefined
let searchMatches: Match[] = []
let searchIndex = -1
let searchError: string | null = null
/** 上一次发给调用方的大纲指纹，只有它变了才 emit */
let outlineSig = ''
/** 页面 DOM 的代次：重建过一次就 +1，用来判断旧的高亮 Range 是否已经失效 */
let domGen = 0
/** 上一次画高亮时的签名（DOM 代次 + 当前匹配 + 全部匹配），一样就不必重画 */
let searchPaintKey = ''

/* -------------------------------------------------------------------------- */
/* 渲染                                                                        */
/* -------------------------------------------------------------------------- */

function numberingOf(model: DocModel): Map<string, string> {
  return computeNumbering(model.blocks, (block) =>
    block.t === 'textBlock' ? resolved.value.styles[block.kind].numbering : 'none',
  )
}

/** 自动编号前缀长度。每次现算：编号会随编辑实时变，缓存反而容易拿到旧值 */
function prefixLength(blockId: string): number {
  return (numberingOf(doc.value).get(blockId) ?? '').length
}

/** 每条批注的锚定文字（Word 叫 scope）。侧栏要显示「批的哪句话」就得靠它 */
const scopes = computed(() => commentScopes(viewDoc.value))

/** 侧栏内容取自渲染快照，不是编辑中的模型 —— 这样打字不会带动侧栏重渲染 */
const comments = computed(() =>
  viewDoc.value.comments
    .filter((c) => c.parentId === undefined)
    .map((c) => ({
      id: c.id,
      author: c.author,
      date: c.date,
      text: c.text,
      scope: scopes.value.get(c.id) ?? '',
      resolved: c.resolved === true,
      replies: viewDoc.value.comments.filter((r) => r.parentId === c.id),
    })),
)

const blocksById = computed(() => {
  const map = new Map<string, TextBlock>()
  for (const block of viewDoc.value.blocks) {
    if (block.t === 'textBlock') map.set(block.id, block)
  }
  return map
})

/** 表格查找表。表格片段没有 data-block-id（见 render/html.ts），得另有一张按 id 取块的表 */
const tablesById = computed(() => {
  const map = new Map<string, TableBlock>()
  for (const block of viewDoc.value.blocks) {
    if (block.t === 'table') map.set(block.id, block)
  }
  return map
})

/**
 * 渲染一个分页片段。
 *
 * 片段偏移用的是「显示文字」坐标系（含自动编号前缀），模型里不含前缀，
 * 所以要先把前缀长度扣掉；只有首片才带前缀。
 *
 * 表格片段走另一条路：按行区间调 renderTableFragment（与量测共用同一个函数），
 * 它渲染出来的格内 div 自己挂 data-block-id，外层不挂。
 */
function fragmentHtml(frag: PageFragment): string {
  if (frag.rowFrom !== undefined) {
    const table = tablesById.value.get(frag.blockId)
    if (!table) return ''
    return renderTableFragment(table, frag.rowFrom, frag.rowTo ?? frag.rowFrom + 1)
  }
  const block = blocksById.value.get(frag.blockId)
  if (!block) return ''
  const prefix = viewNumbering.value.get(frag.blockId) ?? ''
  const p = prefix.length
  const active = activeCommentId.value

  if (frag.from < p) {
    return renderInlinesHtml(
      sliceInlines(block.inlines, 0, Math.max(0, frag.to - p)),
      prefix,
      active,
    )
  }
  return renderInlinesHtml(sliceInlines(block.inlines, frag.from - p, frag.to - p), '', active)
}

/**
 * v-for 的 key。表格片段在同一页里 from/to 恒为 0，同一页出现两张表就会撞键，
 * 所以表格片段必须带上行区间。段落片段也统一带上行区间（没有就是 -1），键保持唯一且稳定。
 */
function fragmentKey(frag: PageFragment): string {
  return `${frag.blockId}:${frag.from}:${frag.rowFrom ?? -1}:${frag.rowTo ?? -1}`
}

function fragmentClass(frag: PageFragment): string {
  return frag.rowFrom !== undefined ? 'wtp-tableFrag' : `wtp-${frag.kind}`
}

/** 批注时间只显示到分钟，够用且不挤 */
function formatDate(iso: string): string {
  const d = new Date(iso)
  if (Number.isNaN(d.getTime())) return ''
  const pad = (n: number): string => String(n).padStart(2, '0')
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`
}

/** 页首那一块不能带段前距，续排块还要去掉首行缩进 —— 这两点分页时已经按此记账 */
function fragmentStyle(frag: PageFragment, isFirst: boolean): Record<string, string> {
  // 表格片段的排版全在表格自己的 CSS 里（行高最小值、单元格内边距），外层不加任何间距
  if (frag.rowFrom !== undefined) return {}
  const style: Record<string, string> = {}
  if (isFirst) style.marginTop = '0'
  if (frag.continuation) style.textIndent = '0'
  return style
}

/**
 * 一节的页面几何（纸宽高 + 页边距 + 方向）。纸宽/纸高/页边距（padding）逐节不同
 * （页面方向按节），所以不能写死在 CSS 里（W5 改动）。
 *
 * 打印用的命名页**只给横排节挂**（`page: wtp-landscape`）。纵排走**默认的 `@page`**
 * —— 它本来就是规格表的纵向尺寸，两者等价；而给纵排也挂 `page: wtp-portrait` 会
 * 让渲染器在最后一页之后从命名页切回默认页，那个切换**强制断页**，于是每份纵排文档
 * 打印出来都多一张空白纸（横竖混排的文档里，横排节结束处同理，这是必要的代价）。
 */
function styleOfSection(section: ResolvedSection | undefined): Record<string, string> {
  const spec = resolved.value
  const size = section?.page.size ?? spec.page.size
  const margin = section?.page.margin ?? spec.page.margin
  const style: Record<string, string> = {
    width: size.width,
    height: size.height,
    padding: `${margin.top} ${margin.right} ${margin.bottom} ${margin.left}`,
  }
  if (section?.settings.orientation === 'landscape') style.page = 'wtp-landscape'
  return style
}

/** 某一页的几何：按它所属的节取 */
function pageStyle(page: PageLayout): Record<string, string> {
  return styleOfSection(sectionOf(page))
}

/** 某一页所属的节（下表越界时兜底到首节） */
function sectionOf(page: PageLayout): ResolvedSection | undefined {
  return resolvedSections.value[page.sectionIndex] ?? resolvedSections.value[0]
}

/** 页码元素的定位：页脚距也是逐节的（横排节的页脚距与纵排节可以不同） */
function pageNumberStyle(page: PageLayout): Record<string, string> {
  return { bottom: (sectionOf(page)?.page ?? resolved.value.page).footer }
}

/** 还没有分页结果时占位那张纸的「页」（几何仍按首节，免得闪一下默认 A4） */
const EMPTY_PAGE: PageLayout = {
  sectionIndex: 0,
  pageNumber: 1,
  showPageNumber: true,
  fragments: [],
  breaks: [],
}

/* -------------------------------------------------------------------------- */
/* 重排                                                                        */
/* -------------------------------------------------------------------------- */

/** 一节的页面几何（纸宽高 + 页边距 + 方向 + 页脚距）—— 模板里直接吃的那几样 */
function geometryOf(section: ResolvedSection): unknown {
  return [styleOfSection(section), section.page.footer]
}

/** 逐节几何的指纹。分页片段一样但几何变了（空节/单行节切方向）也必须重建 DOM */
function geometryKey(sections: readonly ResolvedSection[]): string {
  return JSON.stringify(sections.map(geometryOf))
}

/**
 * 判断这次分页结果与当前渲染的是不是同一份「版面」。
 *
 * 不只看页码与片段：**逐节几何也要比**。反例很实在 —— 把一节改成横排，若那一节的内容
 * 在横竖两种宽度下都只占一行，量测值与分页片段会一模一样，只比片段就会认定「没变化」，
 * 于是页面几何不更新、纸还是竖的（模型却已经是横的，预览与模型就此分家）。
 */
function sameLayout(
  a: readonly PageLayout[],
  b: readonly PageLayout[],
  aSections: readonly ResolvedSection[],
  bSections: readonly ResolvedSection[],
): boolean {
  if (a.length !== b.length) return false
  if (geometryKey(aSections) !== geometryKey(bSections)) return false
  for (let i = 0; i < a.length; i += 1) {
    const pa = a[i]
    const pb = b[i]
    if (!pa || !pb) return false
    if (pa.sectionIndex !== pb.sectionIndex || pa.pageNumber !== pb.pageNumber) return false
    // 页码显示与否也逐节变（关联前节 / 关掉页码），漏比会让「切了开关却没重建 DOM」
    if (pa.showPageNumber !== pb.showPageNumber) return false
    // 换页标记也要比：插一个「文末分页符」不会改变片段，但标记得画出来；
    // 换了是哪几条（kind/blockId）标记的文案与删除目标也不同
    if (pa.breaks.length !== pb.breaks.length) return false
    for (let j = 0; j < pa.breaks.length; j += 1) {
      const ma = pa.breaks[j]
      const mb = pb.breaks[j]
      if (!ma || !mb || ma.blockId !== mb.blockId || ma.kind !== mb.kind) return false
    }
    if (pa.fragments.length !== pb.fragments.length) return false
    for (let j = 0; j < pa.fragments.length; j += 1) {
      const fa = pa.fragments[j]
      const fb = pb.fragments[j]
      if (!fa || !fb) return false
      if (
        fa.blockId !== fb.blockId ||
        fa.from !== fb.from ||
        fa.to !== fb.to ||
        fa.kind !== fb.kind ||
        fa.continuation !== fb.continuation ||
        // 表格片段的行区间也要比：漏比会导致「分页变了却不重建 DOM」，
        // 页面上的表还是上一轮的若干行
        fa.rowFrom !== fb.rowFrom ||
        fa.rowTo !== fb.rowTo
      ) {
        return false
      }
    }
  }
  return true
}

interface RefreshOptions {
  /** 重排后把插入符放回哪里（显示坐标） */
  anchor?: DisplayPoint | null
  /** 选区另一头；给了就用区间还原，不给就还原成插入符 */
  anchorEnd?: DisplayPoint | null
  /** 即使分页没变也重建 DOM（格式化、批注这类不改行数但要改外观的操作） */
  force?: boolean
  /**
   * 锚点落在软换行处时，落到换行**之后**（只有单元格里的 Shift+Enter 用）。
   * 软换行零宽，默认的 placeCaret 会还原到换行之前 —— 那会让回车后敲的字打回上一行。
   */
  afterBreak?: boolean
}

function ensureCss(spec: Spec): void {
  const key = JSON.stringify(spec)
  if (key === cssKey) return
  injectCss(spec)
  cssKey = key
  // 样式换了，量测值全部作废
  clearMeasureCache(cache)
}

/**
 * 量测 → 分页 → （需要时）重建 DOM → 还原插入符。
 *
 * 分页没变就直接返回：DOM 原样不动，插入符天然还在原处，
 * 这正是「正常输入不丢光标」的落点。
 */
function refreshLayout(options: RefreshOptions = {}): void {
  const current = ++token
  const spec = resolved.value
  ensureCss(spec)

  const el = host.value
  if (!el) return

  const items = measureDocument(doc.value, spec, el, cache)
  measured.value = items
  // 逐节解析必须在量测之后、分页之前：量测要按节切探针宽度，分页要逐节版心高
  const sections = resolveSections(doc.value, spec)
  const nextPages = paginate(items, {
    contentHeight: contentBoxPx(spec).height,
    sections: sections.map((s) => ({
      contentHeight: s.content.height,
      showPageNumber: s.showPageNumber,
      restartAtOne: s.settings.restartAtOne,
    })),
  })
  const changed =
    options.force === true || !sameLayout(pages.value, nextPages, resolvedSections.value, sections)

  if (changed) {
    viewDoc.value = cloneDoc(doc.value)
    viewNumbering.value = numberingOf(viewDoc.value)
    pages.value = nextPages
    // 页面几何（纸宽高、页边距、方向）随 viewDoc 一起换新 —— 模板里按节取
    resolvedSections.value = sections
    domGen += 1
    emit('paginated', nextPages.length)
  }

  /*
   * 大纲与查找都必须在快照更新之后重算，且只在数字/指纹真的变了时才 emit：
   * 正常打字不能反复重渲染左栏与计数器（那正是「正常输入不得重排」的反面）。
   * 分页没变时 viewDoc 不动，大纲指纹也就不会变；查找跑在 doc.value 上，
   * 但结果只在有会话时才重算。
   */
  syncOutline()
  if (searchOn) recomputeSearch(false)

  if (!changed) return

  const anchor = options.anchor
  const anchorEnd = options.anchorEnd
  const afterBreak = options.afterBreak === true
  /*
   * 重排之后（DOM 已被 Vue 打过补丁）必须先把片段坐标按分页结果重写一遍，
   * 再还原插入符 —— placeCaret 走的正是这些坐标。见 applyFragmentRanges 的说明。
   */
  void nextTick(() => {
    if (current !== token) return
    const rootEl = root.value
    if (!rootEl) return
    applyFragmentRanges()
    if (anchor) {
      if (
        anchorEnd &&
        (anchorEnd.blockId !== anchor.blockId || anchorEnd.offset !== anchor.offset)
      ) {
        placeRange(rootEl, anchor, anchorEnd)
      } else if (afterBreak) {
        placeCaretAfterBreak(rootEl, anchor)
      } else {
        placeCaret(rootEl, anchor)
      }
    }
    /*
     * 重排会重建片段 DOM，整格复选的高亮得重画（类名挂在 <td> 上，跟着 DOM 一起没了）。
     * 顺手把「重排之后那张表已经不在页面上」的复选收掉（PLAN 13.2 列的退出条件之一）——
     * 一格都不剩（行列被删掉、撤销回到更小的表）时同理；这时还要报一次，
     * 否则调用方的「已选 N 格」会停在旧值上。
     */
    if (cellSelection && cellSelectionCells().length === 0) {
      dropCellSelection()
      emitSelection()
    } else {
      paintCellSelection()
    }
  })
}

async function repaginate(): Promise<void> {
  refreshLayout({ force: true })
}

/* -------------------------------------------------------------------------- */
/* DOM → 模型                                                                  */
/* -------------------------------------------------------------------------- */

/** 片段对应的模型区间（模型文字坐标，不含自动编号前缀） */
function fragmentRange(frag: HTMLElement, blockId: string): { from: number; to: number } {
  const p = prefixLength(blockId)
  const raw = Number(frag.dataset.from ?? '0')
  const end = Number(frag.dataset.to ?? '0')
  const from = Math.max(0, raw - p)
  return { from, to: Math.max(from, end - p) }
}

/**
 * 按当前 DOM 重新给同一容器的各片段打标（data-from / data-to），编辑后坐标才不会越用越偏。
 *
 * 只对挂着 data-block-id 的元素工作，所以它天然只碰段落片段与**格子**：
 * 表格片段的外层没有 data-block-id（见 render/html.ts），不会被当成片段重写属性。
 * 格子永远是一格一个整体（不分页切分），这条路径对它就是把 data-from/to 重写成 0..len。
 *
 * 它写的是「这一眼看上去的 DOM」的坐标，只在下次重排之前有效；重排之后由
 * applyFragmentRanges 按分页结果重写一遍 —— 那时分页结果才是权威。
 */
function retagFragments(blockId: string): void {
  const rootEl = root.value
  if (!rootEl) return
  const frags = Array.from(
    rootEl.querySelectorAll<HTMLElement>(`[data-block-id="${CSS.escape(blockId)}"]`),
  )
  let cursor = 0
  for (const frag of frags) {
    // 直接用 textContent 的长度：自动编号（.wtp-num）就在这一片里，已经算进去了。
    // 早先这里另外又加了一次前缀长度，于是编号段落跨页时坐标整体偏大（表现为丢字）。
    const text = (frag.textContent ?? '').length
    frag.dataset.from = String(cursor)
    frag.dataset.to = String(cursor + text)
    cursor += text
  }
}

/**
 * 按分页结果重写各片段的 data-from / data-to —— **分页结果是这些属性的唯一权威**。
 *
 * 为什么重排之后必须补这一下：`retagFragments` 在每次编辑后命令式改写属性（它按 DOM 现算，
 * 值是「编辑后、重排前」的），而 Vue 打补丁时拿新 vnode 与**上一个 vnode** 比、相等就跳过
 * 属性写入 —— 它不知道 DOM 已经被 retag 改过。两边一旦不一致（典型是「上一页末行让出一个
 * 字位、下一页首字顶上来」，分页给的 from/to 原地不动），DOM 文字与坐标就差一个字符，
 * 之后每次读回都把多出的字符复制进模型，越删越多。见 issues/20260915.md。
 *
 * 只写段落片段：格子的坐标是「格内文字」的 0..len，与分页无关（表格片段外层根本没有
 * data-block-id，格子由 renderTableFragment 每次整段重渲染）。
 */
function applyFragmentRanges(): void {
  const rootEl = root.value
  if (!rootEl) return
  const want = new Map<string, Array<{ from: number; to: number }>>()
  for (const page of pages.value) {
    for (const frag of page.fragments) {
      if (frag.rowFrom !== undefined) continue
      const list = want.get(frag.blockId) ?? []
      list.push({ from: frag.from, to: frag.to })
      want.set(frag.blockId, list)
    }
  }
  // pages 里的片段与 DOM 里的元素都是文档顺序，逐块对位即可
  for (const [blockId, list] of want) {
    const els = rootEl.querySelectorAll<HTMLElement>(`[data-block-id="${CSS.escape(blockId)}"]`)
    for (let i = 0; i < els.length && i < list.length; i += 1) {
      const el = els[i]
      const frag = list[i]
      if (!el || !frag) continue
      el.dataset.from = String(frag.from)
      el.dataset.to = String(frag.to)
    }
  }
}

/**
 * 把页面上每一片的 DOM 读回模型。
 * 片段覆盖的区间用它自己的 data-from/data-to（分页时算出来的行边界），
 * 用 DOM 的新内容整段替换 —— 这样「删掉几个字」和「多打几个字」都能被如实记录。
 *
 * 查找容器而不是段落：格内（data-block-id = cellId）也走这条路，所以格内能直接打字。
 */
function syncPlain(pageEl: HTMLElement): void {
  for (const frag of Array.from(pageEl.querySelectorAll<HTMLElement>('[data-block-id]'))) {
    const blockId = frag.dataset.blockId ?? ''
    const container = findContainer(doc.value, blockId)
    if (!container) continue
    const { from, to } = fragmentRange(frag, blockId)
    replaceRange(container, from, to, readInlines(frag))
    retagFragments(blockId)
  }
}

/** 修订标记；id 在全篇唯一（Word 要求 w:ins / w:del 各自带唯一 id） */
let revSeq = -1
function mark(kind: 'ins' | 'del'): RevMark {
  let max = -1
  // 必须扫到表格格子：修订落在格子里时，只遍历 blocks 里看得见的 textBlock 会漏掉它，
  // 于是下一枚修订标记会重用同一个 id —— Word 要求 w:ins / w:del 的 id 全篇唯一。
  for (const holder of allInlineHolders(doc.value)) {
    for (const inline of holder.inlines) {
      if (inline.t === 'text' && inline.rev) max = Math.max(max, inline.rev.id)
    }
  }
  revSeq = Math.max(revSeq, max) + 1
  return { kind, id: revSeq, author: props.author, date: new Date().toISOString() }
}

/**
 * 修订模式的读回：把这次编辑和模型比一遍。
 * 新增的文字标 w:ins；被删掉的字不真删，原地补回来标 w:del
 * —— 这就是 Word 开着修订时看到的样子。
 */
function syncTracked(pageEl: HTMLElement): DisplayPoint | null {
  let caret: DisplayPoint | null = null

  for (const frag of Array.from(pageEl.querySelectorAll<HTMLElement>('[data-block-id]'))) {
    const blockId = frag.dataset.blockId ?? ''
    const container = findContainer(doc.value, blockId)
    if (!container) continue
    const { from, to } = fragmentRange(frag, blockId)
    const before = sliceStrict(container.inlines, from, to)
    const after = readInlines(frag)

    const oldText = textOf(before)
    const newText = textOf(after)
    if (oldText !== newText) {
      let head = 0
      while (head < oldText.length && head < newText.length && oldText[head] === newText[head]) {
        head += 1
      }
      let tail = 0
      while (
        tail < oldText.length - head &&
        tail < newText.length - head &&
        oldText[oldText.length - 1 - tail] === newText[newText.length - 1 - tail]
      ) {
        tail += 1
      }
      const removedLen = oldText.length - head - tail
      const addedLen = newText.length - head - tail
      const removed: Inline[] = sliceStrict(before, head, head + removedLen).map((inl) =>
        inl.t === 'text' ? { ...inl, rev: mark('del') } : inl,
      )
      const added: Inline[] = sliceStrict(after, head, head + addedLen).map((inl) =>
        inl.t === 'text' ? { ...inl, rev: mark('ins') } : inl,
      )
      replaceRange(container, from, to, [
        ...sliceStrict(before, 0, head),
        ...removed,
        ...added,
        ...sliceStrict(before, head + removedLen, oldText.length),
      ])
      retagFragments(blockId)
      // 落点放在改动之后：删掉的文字之后、或插入的文字之后
      const local = head + removedLen + addedLen
      caret = {
        blockId,
        offset: Number(frag.dataset.from ?? '0') + prefixLengthOf(frag) + local,
      }
    }
  }

  return caret
}

function textOf(inlines: readonly Inline[]): string {
  let out = ''
  for (const inline of inlines) {
    if (inline.t === 'text') out += inline.text
  }
  return out
}

/* -------------------------------------------------------------------------- */
/* 编辑动作                                                                    */
/* -------------------------------------------------------------------------- */

interface Snapshot {
  doc: DocModel
  caret: DisplayPoint | null
  end: DisplayPoint | null
}

const undoStack: Snapshot[] = []
const redoStack: Snapshot[] = []

/** 记一步撤销。caret 给定时用它当「改之前」的落点（结构性操作用得上） */
function pushHistory(caret?: DisplayPoint | null): void {
  undoStack.push({
    doc: cloneDoc(doc.value),
    caret: caret ?? preEditCaret ?? lastCaret,
    end: null,
  })
  if (undoStack.length > 200) undoStack.shift()
  redoStack.length = 0
  preEditCaret = null
}

function selectionRange(): { start: DisplayPoint; end: DisplayPoint } | null {
  const rootEl = root.value
  if (!rootEl) return null
  return currentRange(rootEl)
}

/** 当前落点（优先取实时选区，取不到就用最近一次记录） */
function caretPoint(): DisplayPoint | null {
  return selectionRange()?.start ?? lastCaret
}

/**
 * 焦点已经离开正文时的选区兜底（点下拉切模板、点工具栏都会发生）。
 * 浏览器可能把原生选区整个收掉，那时 selectionRange() 就什么都没有了，
 * 而 stickyRanges 里存着最近一次「真的选中了文字」的区间。
 */
function rangeOfSticky(): { start: DisplayPoint; end: DisplayPoint } | null {
  const first = stickyRanges[0]
  const last = stickyRanges[stickyRanges.length - 1]
  if (!first || !last) return null
  return {
    start: { blockId: first.blockId, offset: first.from },
    end: { blockId: last.blockId, offset: last.to },
  }
}

/** 事件目标所在的页面版心。事件是在 .wtp-content 上派发的，所以目标本身可能不是片段 */
function pageElementOf(node: Node | null): HTMLElement | null {
  if (!node) return null
  const frag = fragmentOf(node)
  const el = frag ?? (node.nodeType === 1 ? (node as Element) : node.parentElement)
  return (el?.closest('.wtp-content') as HTMLElement | null) ?? null
}

function emitSelection(): void {
  if (!props.editable) {
    emit('selection-change', null)
    return
  }
  const rootEl = root.value
  if (!rootEl) return
  const summary = cellSelectionSummary()
  const range = currentRange(rootEl)
  /*
   * 落点：原生选区在正文里就用它。取不到时（整格刷选把原生选区塌掉了、点完按钮焦点走了）
   * 只要还有整格复选，就用复选里的第一格兜底 —— 不兜这一下，「已选 4 格」与对齐/样式按钮的
   * active 态会跟高亮一起消失。
   */
  const fallback = summary && cellSelectionCells()[0]
  const blockId =
    range && findContainer(doc.value, range.start.blockId)
      ? range.start.blockId
      : summary && fallback
        ? cellId(summary.tableId, fallback.row, fallback.col)
        : null
  if (!blockId) {
    emit('selection-change', null)
    return
  }
  const container = findContainer(doc.value, blockId)
  if (!container) {
    emit('selection-change', null)
    return
  }
  // 格子的样式回真实值（缺省 listItem）；段落回它自己的 kind
  const block = findBlock(doc.value, blockId)
  const kind: BlockKind = block ? block.kind : (findCell(doc.value, blockId)?.kind ?? 'listItem')
  const p = prefixLength(blockId)
  const from = range ? Math.max(0, range.start.offset - p) : 0
  const to = range ? Math.max(from, range.end.offset - p) : from
  // 光标在格子里时把表格上下文一并带出去：App 没有响应式的模型，上下文工具条只能靠这一次 emit
  const cell = parseCellId(blockId)
  const table = cell ? findTable(doc.value, cell.tableId) : undefined
  const ctx = cell && table ? tableContextOf(table, cell.row, cell.col) : undefined
  if (ctx && summary) {
    // 有整格复选时，两组对齐的 active 态按**整批的一致性**回显（不一致 → 没有值 → 按钮都不亮）
    ctx.alignH = summary.alignH
    ctx.alignV = summary.alignV
  }
  emit('selection-change', {
    blockId,
    kind,
    from,
    to,
    collapsed: to === from,
    bold: to > from ? rangeIsBold(container, from, to) : false,
    underline: to > from ? rangeIsUnderline(container, from, to) : false,
    color: to > from ? rangeColor(container, from, to) : undefined,
    // 插入符（没选中文字）也要报：点一下修订文字就能整串接受/拒绝
    revisions: hasRevisions(container, from, to),
    ...(ctx ? { table: ctx } : {}),
    ...(summary ? { cellSelection: summary } : {}),
    section: sectionContextOf(blockId),
  })
}

/**
 * 光标所在节的上下文（「节」工具条回显与置灰用）。
 *
 * 三个开关给的是**已解析值**，不是模型里的原始字段 —— 工具条的 radio 选中态
 * 要的是「现在实际是什么」，缺省与继承都在这里算完。
 */
function sectionContextOf(blockId: string): SectionSelectionContext {
  const sections = resolveSections(doc.value, resolved.value)
  const total = sections.length
  const at = Math.min(Math.max(0, sectionIndexOf(doc.value, blockId)), total - 1)
  const settings = sections[at]?.settings
  return {
    index: at,
    total,
    isFirst: at === 0,
    orientation: settings?.orientation ?? 'portrait',
    pageNumbers: settings?.pageNumbers ?? true,
    linkPrevious: settings?.linkPrevious ?? at > 0,
    restartAtOne: settings?.restartAtOne ?? false,
  }
}

/** 表格上下文（工具条回显与禁用用）：行列下标、role、以及**已解析默认值**的两组对齐 */
function tableContextOf(
  table: TableBlock,
  row: number,
  col: number,
): NonNullable<EditorSelection['table']> {
  const cellModel = findCell(doc.value, cellId(table.id, row, col))
  const role = table.rows[row]?.role ?? 'body'
  const styleAlign = resolved.value.styles[cellModel?.kind ?? 'listItem'].align
  return {
    tableId: table.id,
    row,
    col,
    role,
    rows: table.rows.length,
    columns: table.columns,
    bodyRows: table.rows.filter((r) => r.role === 'body').length,
    minLines: table.minLines,
    hasUnit: table.rows.some((r) => r.role === 'unit'),
    hasNote: table.rows.some((r) => r.role === 'note'),
    alignH: cellModel?.align?.h ?? defaultCellAlignH(role, styleAlign),
    alignV: cellModel?.align?.v ?? 'top',
  }
}

function onSelectionChange(): void {
  if (!props.editable) {
    lastCaret = null
    stickyRanges = []
    return
  }
  const rootEl = root.value
  if (!rootEl) return
  /*
   * 整格刷选期间一律不当「落点变更」处理：那会儿原生选区是我们自己塌掉的，
   * 浏览器沿着指针扩出来的区间不是用户意图 —— 它还会把「光标出了这张表」判成真，
   * 于是刷到表格边缘时复选会自己消失。
   */
  if (cellDrag?.brushing) return
  const sel = document.getSelection()
  if (!sel || sel.rangeCount === 0) return
  const range = sel.getRangeAt(0)
  if (!fragmentOf(range.startContainer)) return
  // 选区回到正文了，那层「续命」高亮就该撤掉，免得和原生选区叠成两种颜色
  dropKeptSelection()
  const selected = selectedRanges(rootEl)
  if (selected.length > 0) stickyRanges = selected
  const point = displayPointOf(rootEl, range.startContainer, range.startOffset)
  if (point && range.collapsed) lastCaret = point
  // 光标移出那张表（点正文别处、方向键走到表外、切模板重排）→ 复选收起
  const caretCell = parseCellId(point?.blockId ?? '')
  if (cellSelection && (!caretCell || caretCell.tableId !== cellSelection.tableId)) {
    dropCellSelection()
  }
  emitSelection()
}

/** 读回 + 重排；onInput 与组字结束兜底共用 */
function runSync(pageEl: HTMLElement): void {
  pushHistory()
  let anchor: DisplayPoint | null = null
  if (props.trackChanges) anchor = syncTracked(pageEl)
  else syncPlain(pageEl)
  if (!anchor) anchor = selectionRange()?.start ?? lastCaret
  refreshLayout({ anchor, force: props.trackChanges })
}

function onInput(event: Event): void {
  if (!props.editable || composing) return
  if (compSyncTimer !== null) {
    clearTimeout(compSyncTimer)
    compSyncTimer = null
  }
  const pageEl = pageElementOf(event.target as Node)
  if (!pageEl) return
  runSync(pageEl)
}

/** 选区碰到的片段元素（文档顺序）。判断「这次删除会不会把多个片段并到一起」用它 */
function fragmentsInRange(rootEl: HTMLElement, range: Range): HTMLElement[] {
  return Array.from(rootEl.querySelectorAll<HTMLElement>('[data-block-id]')).filter((el) =>
    range.intersectsNode(el),
  )
}

/**
 * 选区端点的显示坐标。端点落在片段之外时（整页全选时端点就是 `.wtp-content` 本身）
 * 退到最外侧被碰到的片段那一端，这样「全选本页再删」也能算出一个完整区间。
 */
function endpointOf(
  rootEl: HTMLElement,
  touched: readonly HTMLElement[],
  node: Node,
  offset: number,
  side: 'start' | 'end',
): DisplayPoint | null {
  const direct = displayPointOf(rootEl, node, offset)
  if (direct) return direct
  const frag = side === 'start' ? touched[0] : touched[touched.length - 1]
  const blockId = frag?.dataset.blockId ?? ''
  if (!frag || blockId === '') return null
  const start = Number(frag.dataset.from ?? '0')
  return {
    blockId,
    offset: side === 'start' ? start : start + (frag.textContent ?? '').length,
  }
}

/**
 * 插入符所在片段在显示坐标里的两端（判断是否顶到片段的左／右边缘用）。
 *
 * 片段按**插入符所在的 DOM 节点**找（fragmentOf 往上取），不能用按数值区间找的
 * `fragmentAt`：相邻两片共用一个边界（前片的 to == 后片的 from）时，数值法会命中前一片，
 * 于是「顶在下一页首字之前」被误判成「在上一页末尾之中」，跨页那一退就接不了管。
 */
function fragmentEdges(): { from: number; to: number } | null {
  const rootEl = root.value
  if (!rootEl) return null
  const sel = document.getSelection()
  if (!sel || sel.rangeCount === 0) return null
  const frag = fragmentOf(sel.getRangeAt(0).startContainer)
  if (!frag) return null
  const from = Number(frag.dataset.from ?? '0')
  return { from, to: from + (frag.textContent ?? '').length }
}

/**
 * 这次选区删除要不要由模型层接管。
 *
 * 单独拆出来是为了让调用方**先问再改**：`deleteSelection` 一进去就动模型，而撤销快照
 * 必须在动模型之前记 —— 先问清楚，才不会给「交给浏览器做」的情况白记一步撤销。
 */
function selectionNeedsTakeover(range: Range, always = false): boolean {
  const rootEl = root.value
  if (!rootEl || range.collapsed) return false
  const touched = fragmentsInRange(rootEl, range).length
  return always ? touched > 0 : touched > 1
}

/**
 * 删掉选中的文字，返回删完后的落点（显示坐标）。
 *
 * **选区碰到两个以上片段时必须由模型层做，不能放给浏览器**：原生删除会把相邻的片段元素
 * 并成一个、把其余的直接删掉，而 Vue 手里还留着那些节点的 vnode，页面就再也补不回来
 * （issues/20260915.md 的「全选删除本页，下一页文字没有前移」）。
 *
 * 段落之间按 Word 语义接起来（deleteSpan：段落标记被删掉，两头并成一段）；端点落在表格
 * 格子里时退回「各容器各自删掉选中的文字」—— 格子的单段落模型表达不了并格。
 *
 * 只碰到一个片段时默认不接管（片内删除交给浏览器，原生手感最好，读回也只有那一片），
 * 返回 null；`always` 为真时连片内选区也自己做 —— 粘贴要替换选区，而粘贴是调用方
 * 自己 preventDefault 的，浏览器那条路已经没了。
 * **调用方要在它之前 `pushHistory`**（见 selectionNeedsTakeover 的说明）。
 */
function deleteSelection(range: Range, always = false): DisplayPoint | null {
  const rootEl = root.value
  if (!rootEl || range.collapsed) return null
  const touched = fragmentsInRange(rootEl, range)
  if (touched.length === 0) return null
  if (touched.length < 2 && !always) return null
  const start = endpointOf(rootEl, touched, range.startContainer, range.startOffset, 'start')
  const end = endpointOf(rootEl, touched, range.endContainer, range.endOffset, 'end')
  if (!start || !end) return null
  // 落点先落到**模型坐标**上：删除会把段落并掉、连带改了自动编号前缀，
  // 显示坐标要按删除后的前缀重算一次
  const caret: BlockPoint = {
    blockId: start.blockId,
    offset: Math.max(0, start.offset - prefixLength(start.blockId)),
  }
  const rev = props.trackChanges ? mark('del') : undefined
  const head = findBlock(doc.value, start.blockId)
  const tail = findBlock(doc.value, end.blockId)

  if (head && tail) {
    deleteSpan(
      doc.value,
      caret,
      { blockId: end.blockId, offset: Math.max(0, end.offset - prefixLength(end.blockId)) },
      rev,
    )
  } else {
    /*
     * 端点落在表格格子上时（选区的一端正好是表格：全选一整页、而这一页末尾摊着一行表格，
     * 就是这种情形），**跨段并段那条 Word 语义不该因此失效** —— 段落的两个端点改取选区里
     * 最外侧的**段落**，格子（表达不了并格）仍旧各按容器删。不这么分一下，整页删除会退化成
     * 「把每一段各自清空」，段落不再并成一段（2026-09-16 W7 实测：② 让样本表的表头行落到
     * 第 1 页之后，X6 的那条断言就红在这里）。
     */
    const ranges = selectedRanges(rootEl)
    const para = (r: DisplayRange): boolean => parseCellId(r.blockId) === null
    const headPara = ranges.find(para)
    const tailPara = [...ranges].reverse().find(para)
    if (headPara && tailPara) {
      const spanStart: BlockPoint = {
        blockId: headPara.blockId,
        offset: Math.max(0, headPara.from - prefixLength(headPara.blockId)),
      }
      deleteSpan(
        doc.value,
        spanStart,
        {
          blockId: tailPara.blockId,
          offset: Math.max(0, tailPara.to - prefixLength(tailPara.blockId)),
        },
        rev,
      )
      caret.blockId = spanStart.blockId
      caret.offset = spanStart.offset
      for (const r of ranges) {
        if (para(r)) continue
        const p = prefixLength(r.blockId)
        deleteRange(doc.value, r.blockId, Math.max(0, r.from - p), Math.max(0, r.to - p), rev)
      }
    } else {
      for (const r of ranges) {
        const p = prefixLength(r.blockId)
        deleteRange(doc.value, r.blockId, Math.max(0, r.from - p), Math.max(0, r.to - p), rev)
      }
    }
  }

  return { blockId: caret.blockId, offset: prefixLength(caret.blockId) + caret.offset }
}

function onBeforeInput(event: InputEvent): void {
  if (!props.editable) return
  const rootEl = root.value
  if (!rootEl) return
  const sel = document.getSelection()
  if (!sel || sel.rangeCount === 0) return
  const range = sel.getRangeAt(0)

  if (!range.collapsed) {
    if (!selectionNeedsTakeover(range)) {
      // 片内选区：交给浏览器（原生手感、原生的撤销分组都在它那边）
      preEditCaret = displayPointOf(rootEl, range.startContainer, range.startOffset)
      return
    }
    event.preventDefault()
    const before = startPointOf(range) ?? lastCaret
    pushHistory(before)
    const caret = deleteSelection(range) ?? before
    // 端点算不出来时（极少见）这次删除不做；上面那步撤销是白记的，无害
    if (!caret) return
    /*
     * 这次的输入类型本身就是「插入」（打字、输入法、粘贴）时，选区删完还得把内容补上
     * —— 数据只有 `event.data` 里有，拿不到就先只删（这几条极少见，也不该把 DOM 并坏）。
     */
    const data = event.data ?? ''
    let anchor = caret
    if (event.inputType === 'insertText' && data !== '') {
      const offset = Math.max(0, caret.offset - prefixLength(caret.blockId))
      insertText(doc.value, caret.blockId, offset, data, props.trackChanges ? mark('ins') : undefined)
      anchor = { blockId: caret.blockId, offset: caret.offset + data.length }
    }
    refreshLayout({ anchor, force: true })
    void nextTick(emitSelection)
    return
  }

  if (!fragmentOf(range.startContainer)) {
    // 落点在正文下方空白处：拉回最后一段末尾，免得敲进来的字没有归宿
    const frags = Array.from(rootEl.querySelectorAll<HTMLElement>('[data-block-id]'))
    const last = frags[frags.length - 1]
    if (last) {
      event.preventDefault()
      placeCaret(rootEl, {
        blockId: last.dataset.blockId ?? '',
        offset: Number(last.dataset.to ?? '0'),
      })
    }
    return
  }
  preEditCaret = displayPointOf(rootEl, range.startContainer, range.startOffset)
}

/** 本次输入之前的落点（编辑器已经在选区里，pushHistory 的「改之前」要用它） */
function startPointOf(range: Range): DisplayPoint | null {
  const rootEl = root.value
  if (!rootEl) return null
  return displayPointOf(rootEl, range.startContainer, range.startOffset)
}

function onCompositionStart(): void {
  composing = true
  /*
   * 打字/输入法落在跨片段选区上：先按模型把选区删掉再让输入法写字。
   * 不接管的话原生会把片段元素并掉（见 deleteSelection 的说明）。
   */
  if (!props.editable) return
  const sel = document.getSelection()
  if (!sel || sel.rangeCount === 0) return
  const range = sel.getRangeAt(0)
  if (!selectionNeedsTakeover(range)) return
  const before = startPointOf(range) ?? lastCaret
  pushHistory(before)
  const caret = deleteSelection(range) ?? before
  if (caret) refreshLayout({ anchor: caret, force: true })
  void nextTick(emitSelection)
}

function onCompositionEnd(): void {
  composing = false
  if (!props.editable) return
  // 多数浏览器在 compositionend 之后还会补一个 input；这里留个兜底，
  // input 先到就把兜底取消，免得同一笔输入读回两次。
  if (compSyncTimer !== null) clearTimeout(compSyncTimer)
  compSyncTimer = setTimeout(() => {
    compSyncTimer = null
    const sel = document.getSelection()
    if (!sel || sel.rangeCount === 0) return
    const pageEl = pageElementOf(sel.getRangeAt(0).startContainer)
    if (pageEl) runSync(pageEl)
  }, 0)
}

/**
 * 这次按键命中哪个动作。表按动作集的固定顺序扫，所以「两个动作绑同一个组合键」时
 * 先到先得（resolveShortcuts 已经把后到的解绑并 warn 过一遍，这里只是照表办事）。
 * Tab 与方向键**不在表里**：它们是编辑器手感（跨格移动），不是可配置的动作。
 */
function shortcutActionOf(event: KeyboardEvent): ShortcutAction | null {
  const table = shortcuts.value
  for (const action of SHORTCUT_ACTIONS) {
    if (matchShortcut(event, table[action])) return action
  }
  return null
}

/* -------------------------------------------------------------------------- */
/* F4：重复上一步                                                              */
/* -------------------------------------------------------------------------- */

/**
 * 最后一次**可重复的格式操作**（不含删除 / 并段 / 插入这类破坏性、结构性的操作 ——
 * 重放它们会把撤销栈搞乱）。记的是「调用本身」，所以 F4 作用的是**当前**落点 / 选区
 * （Word 的 F4 也是这个语义），不是操作发生时的位置。
 */
let repeatAction: (() => boolean) | null = null

/**
 * 记下并执行一次可重复的操作。只有**真的改了模型**（返回 true）才记 —— 否则
 * 「没选中文字时点一下加粗」也会占住 F4，之后按 F4 只会静默空转，
 * 连「没有可重复的操作」这句提示都出不来。
 *
 * 返回值就是「这一步真的改了模型没有」：工具条那条路不看它（按钮该置灰的自己置灰），
 * 键盘那条路（`onKeydown`）拿它决定要不要提示。
 */
function repeatable(run: () => boolean): boolean {
  if (!run()) return false
  repeatAction = run
  return true
}

/** F4 的处理：没有可重复的操作就空转 + 一句提示（与金额格式化失败同一套提示条） */
function repeatLast(): void {
  if (!repeatAction) {
    emit('toast', '没有可重复的操作')
    return
  }
  repeatAction()
}

function onKeydown(event: KeyboardEvent): void {
  if (!props.editable || composing) return

  const mod = event.ctrlKey || event.metaKey
  const action = shortcutActionOf(event)
  if (action) {
    event.preventDefault()
    switch (action) {
      case 'bold':
        toggleBold()
        break
      case 'underline':
        toggleUnderline()
        break
      case 'trackChanges':
        emit('toggle-track-changes')
        break
      // ctrl+F 是浏览器查找、ctrl+G 是「查找下一个」，不 preventDefault 就抢不回来
      case 'find':
        emit('open-search', 'find')
        break
      case 'replace':
        emit('open-search', 'replace')
        break
      case 'undo':
        undo()
        break
      case 'redo':
        redo()
        break
      case 'formatAmount':
        if (!formatSelectionAsAmount()) emit('toast', '选中内容不是有效数字')
        break
      case 'repeat':
        repeatLast()
        break
      /*
       * 下面这 20 条是 W9 追加的（键位全在 Ctrl+Alt+… 一档，理由见 README「快捷键」）。
       * 两条需要反馈的路子：
       *   · 接受 / 拒绝修订：没有覆盖到的修订时一步都不许动，弹一句；
       *   · 格里对齐：判据与「表格」页那六枚按钮的置灰条件同源 —— 光标（或整格复选）
       *     落到某个格子上才动手。这里除了看 setter 的返回值，还要求「真的没有目标格」：
       *     返回值同时覆盖「本来就是这个值、又没有覆盖可清」的空转（按钮那条路是静默的），
       *     单看返回值会在格内弹一句位置判断错误的提示。
       */
      case 'colorRed':
        setColor('FF0000')
        break
      case 'colorClear':
        setColor(null)
        break
      case 'acceptRevision':
        if (!resolveRevisionsOf('accept')) emit('toast', '这里没有可以接受的修订')
        break
      case 'rejectRevision':
        if (!resolveRevisionsOf('reject')) emit('toast', '这里没有可以拒绝的修订')
        break
      case 'styleTitle':
        setBlockKind('title')
        break
      case 'styleH1':
        setBlockKind('h1')
        break
      case 'styleH2':
        setBlockKind('h2')
        break
      case 'styleH3':
        setBlockKind('h3')
        break
      case 'styleBody':
        setBlockKind('body')
        break
      case 'spaceEm':
        insertSpecialSpace('em')
        break
      case 'spaceEn':
        insertSpecialSpace('en')
        break
      case 'spaceQuarterEm':
        insertSpecialSpace('quarterEm')
        break
      case 'breakSection':
        insertSectionBreak()
        break
      case 'breakPage':
        insertPageBreak()
        break
      case 'cellAlignLeft':
        if (!setTableCellAlignH('left') && !cellTargets()) {
          emit('toast', '请先把光标放进表格的格子里')
        }
        break
      case 'cellAlignCenter':
        if (!setTableCellAlignH('center') && !cellTargets()) {
          emit('toast', '请先把光标放进表格的格子里')
        }
        break
      case 'cellAlignRight':
        if (!setTableCellAlignH('right') && !cellTargets()) {
          emit('toast', '请先把光标放进表格的格子里')
        }
        break
      case 'cellAlignTop':
        if (!setTableCellAlignV('top') && !cellTargets()) {
          emit('toast', '请先把光标放进表格的格子里')
        }
        break
      case 'cellAlignMiddle':
        if (!setTableCellAlignV('middle') && !cellTargets()) {
          emit('toast', '请先把光标放进表格的格子里')
        }
        break
      case 'cellAlignBottom':
        if (!setTableCellAlignV('bottom') && !cellTargets()) {
          emit('toast', '请先把光标放进表格的格子里')
        }
        break
    }
    return
  }
  /*
   * ctrl+shift+Z 是「重做」的另一条惯例键（ctrl+Z 的兄弟键）。动作表是「一个动作一个组合键」，
   * 塞不进第二条，所以放在表外兜底：表里没命中才轮到它 —— 调用方要是把某个动作改绑到
   * ctrl+shift+Z，那一条自然优先。必须拦下：放给浏览器会落到原生重做，而本项目不用原生
   * 撤销栈，那一刀会把 DOM 与模型分开。
   */
  if (mod && event.shiftKey && !event.altKey && (event.key === 'z' || event.key === 'Z')) {
    event.preventDefault()
    redo()
    return
  }
  // 整格复选：Esc 收起（PLAN 13.2 的退出条件之一）。没有复选时不接管，Esc 照旧留给别处
  if (event.key === 'Escape' && cellSelection) {
    event.preventDefault()
    clearCellSelection()
    return
  }
  if (event.key === 'Tab') {
    // 编辑公文时 Tab 不该把焦点跳出去；在格内则行优先跨格
    event.preventDefault()
    const point = caretPoint()
    const cell = point ? parseCellId(point.blockId) : null
    const table = cell ? findTable(doc.value, cell.tableId) : undefined
    if (!cell || !table) return
    // 已经是最后一格（Shift+Tab 是第一格）时 stepCell 返回 null → 不作任何响应
    // （用户拍板：不换行、不自动加行、不跳出焦点）
    const step = stepCell(table, cell.row, cell.col, event.shiftKey ? 'prev' : 'next')
    if (step) moveCaretToCell(cell.tableId, step)
    return
  }
  if (event.key === 'Enter') {
    /*
     * 格内 Enter 一律不接管（单元格是单段落，分段表达不出来）—— 但护栏现在就得有：
     * 不拦的话回车会掉进 insertParagraphBreak，把整张表当段落切开。
     * （insertParagraphBreak 里还有一道 findBlock 兜底，两道都在，免得日后改一处漏一处。）
     * 格内 Shift+Enter 插一枚软换行（Word 的 <w:br/>），插入符落到换行之后。
     */
    const point = caretPoint()
    event.preventDefault()
    if (point && isTableCell(point.blockId)) {
      if (event.shiftKey) insertSoftBreak(point)
      return
    }
    insertParagraphBreak()
    return
  }
  if (event.key === 'Backspace' && !mod) {
    const sel = document.getSelection()
    if (!sel || !sel.isCollapsed) return
    const point = caretPoint()
    if (!point) return
    /*
     * 格内退格单独守住：格首那一退如果放给浏览器，原生会把相邻 `<td>` 的 DOM 并掉，
     * 表格结构当场就坏了（格内偏移为 0 = 光标顶在格首；格子的前缀恒为 0）。
     * 格内别的位置不动，交给浏览器在格内正常退格。
     */
    if (isTableCell(point.blockId)) {
      if (point.offset <= prefixLength(point.blockId)) event.preventDefault()
      return
    }
    /*
     * 段中间交给浏览器在片内正常退格。但**顶在片段左缘**时不能放给它：段落被分页切到
     * 下一页的那一片，左缘并不是段首，而浏览器跨不过 contenteditable 的边界
     * （每页一个宿主），这一退在浏览器手里会变成什么都不做（issues/20260915.md）。
     * 这时按模型往前删一个字 —— 上一个字可能在上一页的片段里，模型坐标不管分页。
     */
    if (point.offset > prefixLength(point.blockId)) {
      const edges = fragmentEdges()
      if (!edges || point.offset > edges.from) return
      const offset = point.offset - prefixLength(point.blockId)
      event.preventDefault()
      pushHistory(point)
      deleteRange(
        doc.value,
        point.blockId,
        offset - 1,
        offset,
        props.trackChanges ? mark('del') : undefined,
      )
      refreshLayout({ anchor: { blockId: point.blockId, offset: point.offset - 1 }, force: true })
      void nextTick(emitSelection)
      return
    }
    /*
     * 没得并（首块、或前一块是表格／换页标记）时也要拦：放给浏览器的话，
     * 原生会把相邻的片段元素并起来，Vue 手里的 vnode 却还在，DOM 就补不回来了。
     */
    event.preventDefault()
    if (!canMergeIntoPrevious(doc.value, point.blockId)) return
    // 撤销快照必须记在改模型之前（并完再记就退回不去了）
    pushHistory(point)
    const merged = mergeIntoPrevious(doc.value, point.blockId)
    if (!merged) return
    refreshLayout({
      anchor: {
        blockId: merged.blockId,
        offset: prefixLength(merged.blockId) + merged.offset,
      },
      force: true,
    })
    return
  }
  if (event.key === 'Delete' && !mod) {
    const sel = document.getSelection()
    if (!sel || !sel.isCollapsed) return
    const point = caretPoint()
    if (!point) return
    if (isTableCell(point.blockId)) {
      /*
       * 格尾的 Delete 同样是结构性的：不拦的话原生会跟下一格合并。
       * 只在「格内偏移已到格文字长度」时接管，格内其它位置留给浏览器。
       */
      const cellContainer = findContainer(doc.value, point.blockId)
      if (!cellContainer) return
      if (point.offset >= prefixLength(point.blockId) + containerLength(cellContainer)) {
        event.preventDefault()
      }
      return
    }
    const container = findContainer(doc.value, point.blockId)
    const edges = fragmentEdges()
    if (!container || !edges || point.offset < edges.to) return
    const offset = Math.max(0, point.offset - prefixLength(point.blockId))
    event.preventDefault()
    if (offset < containerLength(container)) {
      // 片段右缘，但这一段后面还有字（跨页时它在下一页那一片里）：按模型删掉那一个字
      pushHistory(point)
      deleteRange(
        doc.value,
        point.blockId,
        offset,
        offset + 1,
        props.trackChanges ? mark('del') : undefined,
      )
      refreshLayout({ anchor: point, force: true })
      void nextTick(emitSelection)
      return
    }
    /*
     * 段尾：删的是「段落标记」，也就是把下一段接上来（Word 语义）。
     * 修订模式下不做（模型里没有段落标记可留痕），但默认行为一律拦住 ——
     * 原生并片段 DOM 会把结构弄坏。下一块是表格／换页标记时不改模型、也不放行。
     */
    if (props.trackChanges) return
    if (!canJoinWithNext(doc.value, point.blockId)) return
    // 撤销快照必须记在改模型之前（并完再记就退回不去了）
    pushHistory(point)
    const joined = joinWithNext(doc.value, point.blockId)
    if (!joined) return
    refreshLayout({
      anchor: { blockId: joined.blockId, offset: prefixLength(joined.blockId) + joined.offset },
      force: true,
    })
    void nextTick(emitSelection)
    return
  }
  /*
   * ← / → 跨格：只在键盘事件不带任何修饰键、且落点在格内时才考虑。
   * 判据是「落点已在格内偏移 0」（←）或「落点已到该格文字末尾」（→），
   * 满足才接管；其余一律 return 走原路，别影响段落里的左右移动。
   */
  if (event.key === 'ArrowLeft' || event.key === 'ArrowRight') {
    if (mod || event.altKey || event.shiftKey) return
    const sel = document.getSelection()
    if (!sel || !sel.isCollapsed) return
    const point = caretPoint()
    const cell = point ? parseCellId(point.blockId) : null
    const table = cell ? findTable(doc.value, cell.tableId) : undefined
    if (!point || !cell || !table) return
    const container = findContainer(doc.value, point.blockId)
    if (!container) return
    const start = prefixLength(point.blockId)
    const atStart = point.offset <= start
    const atEnd = point.offset >= start + containerLength(container)
    const dir = event.key === 'ArrowLeft' ? 'prev' : 'next'
    if (dir === 'prev' ? !atStart : !atEnd) return
    const step = stepCell(table, cell.row, cell.col, dir)
    if (!step) return
    event.preventDefault()
    moveCaretToCell(cell.tableId, step)
    return
  }
  /*
   * ↑ / ↓ 跨格：只有插入符位于该格的首 / 末**视觉行**时才接管，靠 DOM 量测判定。
   * 判不准就 fail-open（不 preventDefault），宁可少管，也不能把普通的上下移动吃掉。
   */
  if (event.key === 'ArrowUp' || event.key === 'ArrowDown') {
    if (mod || event.altKey || event.shiftKey) return
    const sel = document.getSelection()
    if (!sel || !sel.isCollapsed) return
    const point = caretPoint()
    const cell = point ? parseCellId(point.blockId) : null
    const table = cell ? findTable(doc.value, cell.tableId) : undefined
    if (!point || !cell || !table) return
    const step = visualLineStep(table, cell.row, cell.col, event.key === 'ArrowUp' ? 'up' : 'down')
    if (!step) return
    event.preventDefault()
    moveCaretToCell(cell.tableId, step)
  }
}

/**
 * 把插入符落到目标格。跨格**只挪原生选区、不改模型、不 refreshLayout**（DOM 没重建）。
 * 落点偏移按显示坐标算（格子的自动编号前缀恒为 0，仍按既有写法扣一次更稳），
 * 落完补一次 emitSelection，让工具条的落点提示跟上。
 */
function moveCaretToCell(tableId: string, step: CellStep): void {
  const rootEl = root.value
  if (!rootEl) return
  const id = cellId(tableId, step.row, step.col)
  const container = findContainer(doc.value, id)
  if (!container) return
  const offset = prefixLength(id) + (step.at === 'end' ? containerLength(container) : 0)
  placeCaret(rootEl, { blockId: id, offset })
  emitSelection()
}

/** 插入符的 rect top；拿不到（选区不在、rect 退化成长宽都为 0）返回 null —— 调用方据此 fail-open */
function caretRectTop(): number | null {
  const sel = document.getSelection()
  if (!sel || sel.rangeCount === 0) return null
  const range = sel.getRangeAt(0)
  if (!range.collapsed) return null
  // 复制一份 range，setStart/setEnd 都设在同一 (node, offset) 上再取 rect；
  // 不往 DOM 里插临时 span（那会动 DOM）。
  const probe = document.createRange()
  probe.setStart(range.startContainer, range.startOffset)
  probe.setEnd(range.startContainer, range.startOffset)
  const rect = probe.getBoundingClientRect()
  if (rect.width === 0 && rect.height === 0) return null
  return rect.top
}

/** 格内首 / 末字符的 rect top；格内没有文字或 rect 退化时返回 null */
function charRectTop(frag: HTMLElement, which: 'first' | 'last'): number | null {
  const walker = document.createTreeWalker(frag, NodeFilter.SHOW_TEXT)
  let first: Text | null = null
  let last: Text | null = null
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const text = node as Text
    if (text.data === '' || text.parentElement?.closest('.wtp-num')) continue
    if (!first) first = text
    last = text
  }
  const target = which === 'first' ? first : last
  if (!target) return null
  const range = document.createRange()
  if (which === 'first') {
    range.setStart(target, 0)
    range.setEnd(target, 1)
  } else {
    const n = target.data.length
    range.setStart(target, n - 1)
    range.setEnd(target, n)
  }
  const rect = range.getClientRects()[0] ?? range.getBoundingClientRect()
  if (!rect || (rect.width === 0 && rect.height === 0)) return null
  return rect.top
}

/**
 * ↑ / ↓ 的落点判定：插入符 top ≈ 格内首字符 top → 已在首视觉行（↑ 接管、↓ 不接管）；
 * ≈ 末字符 top → 已在末视觉行（↓ 接管、↑ 不接管）；中间视觉行一律不接管。
 * 任何一处取不到（rect 退化、首末字符缺失）或抛异常 → 返回 null（fail-open）。
 */
function visualLineStep(
  table: TableBlock,
  row: number,
  col: number,
  dir: 'up' | 'down',
): CellStep | null {
  try {
    const rootEl = root.value
    if (!rootEl) return null
    const frag = rootEl.querySelector<HTMLElement>(
      `[data-block-id="${CSS.escape(cellId(table.id, row, col))}"]`,
    )
    if (!frag) return null
    const caretTop = caretRectTop()
    const firstTop = charRectTop(frag, 'first')
    const lastTop = charRectTop(frag, 'last')
    if (caretTop === null || firstTop === null || lastTop === null) return null
    const tol = 1
    if (dir === 'up' && Math.abs(caretTop - firstTop) <= tol) {
      return verticalCell(table, row, col, 'up')
    }
    if (dir === 'down' && Math.abs(caretTop - lastTop) <= tol) {
      return verticalCell(table, row, col, 'down')
    }
    return null
  } catch {
    return null
  }
}

/** 单元格里 Shift+Enter：在插入符处插一枚软换行（零宽），插入符落到换行之后 */
function insertSoftBreak(point: DisplayPoint): void {
  const container = findContainer(doc.value, point.blockId)
  if (!container) return
  const offset = Math.max(0, point.offset - prefixLength(point.blockId))
  pushHistory(point)
  replaceRange(container, offset, offset, [{ t: 'break' }])
  refreshLayout({
    // 模型偏移在换行处不变（零宽），靠 afterBreak 把插入符落到换行之后
    anchor: { blockId: point.blockId, offset: prefixLength(point.blockId) + offset },
    force: true,
    afterBreak: true,
  })
}

/** 这个 id 是表格格子（`tableId.rNcM`）还是普通段落？格子的结构性操作都得绕开 */
function isTableCell(blockId: string): boolean {
  return parseCellId(blockId) !== null
}

/** 回车：在落点切开当前块。标题类段落回车后接一个正文段（公文习惯：标题一行一段） */
function insertParagraphBreak(): void {
  const point = caretPoint()
  if (!point) return
  // splitBlock 只对段落有意义；格子落点到这里应当是空操作（onKeydown 已经拦了一道）
  const block = findBlock(doc.value, point.blockId)
  if (!block) return
  const offset = Math.max(0, point.offset - prefixLength(point.blockId))
  const tailKind: BlockKind =
    block.kind === 'body' || block.kind === 'listItem' ? block.kind : 'body'
  pushHistory(point)
  const created = splitBlock(doc.value, block.id, offset, tailKind)
  if (!created) return
  refreshLayout({ anchor: { blockId: created, offset: prefixLength(created) }, force: true })
}

function onPaste(event: ClipboardEvent): void {
  if (!props.editable) return
  event.preventDefault()
  const text = event.clipboardData?.getData('text/plain') ?? ''
  if (text === '') return
  const rootEl = root.value
  const sel = document.getSelection()
  const range = sel && sel.rangeCount > 0 ? sel.getRangeAt(0) : null
  let point = caretPoint()
  if (!point) return
  pushHistory(point)

  /*
   * 粘贴要**替换**选中的文字（Word 语义）。粘贴是我们自己 preventDefault 的，
   * 浏览器那条路已经没了，所以这里必须自己删 —— 跨片段时更不能放给原生（见 deleteSelection）。
   */
  if (rootEl && range && !range.collapsed) {
    point = deleteSelection(range, true) ?? point
  }

  // 纯文本粘贴：换行还原成新段落（公文里段落就是行）
  const lines = text.replace(/\r\n?/g, '\n').split('\n')

  if (isTableCell(point.blockId)) {
    /*
     * 单元格是单段落，分段在这里表达不出来。多行粘贴把换行变成**软换行**
     * （而不是像段落那样往后切段）—— 若不特判，splitBlock 对格子是空操作，
     * 后面的行会被静默丢掉。
     */
    const container = findContainer(doc.value, point.blockId)
    if (!container) return
    const offset = Math.max(0, point.offset - prefixLength(point.blockId))
    const inlines: Inline[] = []
    lines.forEach((line, i) => {
      if (i > 0) inlines.push({ t: 'break' })
      if (line !== '') inlines.push({ t: 'text', text: line })
    })
    replaceRange(container, offset, offset, inlines)
    refreshLayout({
      anchor: {
        blockId: point.blockId,
        offset: prefixLength(point.blockId) + offset + lines.join('').length,
      },
      force: true,
    })
    return
  }

  let blockId = point.blockId
  let offset = Math.max(0, point.offset - prefixLength(blockId))
  for (let i = 0; i < lines.length; i += 1) {
    if (i > 0) {
      const created = splitBlock(doc.value, blockId, offset, 'body')
      if (!created) break
      blockId = created
      offset = 0
    }
    const line = lines[i] ?? ''
    insertText(doc.value, blockId, offset, line)
    offset += line.length
  }
  refreshLayout({
    anchor: { blockId, offset: prefixLength(blockId) + offset },
    force: true,
  })
}

/* -------------------------------------------------------------------------- */
/* 工具栏 API                                                                  */
/* -------------------------------------------------------------------------- */

/** 对当前选区覆盖到的每一块依次施加操作（显示坐标 → 模型坐标在这里换算） */
function forSelection(
  mutate: (model: DocModel, blockId: string, from: number, to: number) => void,
): void {
  const rootEl = root.value
  if (!rootEl) return
  const ranges = selectedRanges(rootEl)
  if (ranges.length === 0) return
  const before = selectionRange()
  pushHistory()
  for (const range of ranges) {
    const p = prefixLength(range.blockId)
    const from = Math.max(0, range.from - p)
    const to = Math.max(from, range.to - p)
    mutate(doc.value, range.blockId, from, to)
  }
  refreshLayout({
    anchor: before?.start ?? null,
    anchorEnd: before?.end ?? null,
    force: true,
  })
}

/**
 * 段落样式 / 格内样式。
 *
 * 格内样式走**目标格列表**（整格复选时是整批；没有复选时是原生选区覆盖到的格子、
 * 或落点那一格）—— 单选时列表就一格，与多选是同一条代码。
 * 段落仍按老路走（选区覆盖到的每一块）。返回值 = 「这一下有没有落到模型上」，供 F4 决定要不要记它。
 */
function applyBlockKind(kind: BlockKind): boolean {
  const rootEl = root.value
  if (!rootEl) return false
  const range = currentRange(rootEl)
  const group = cellTargets()
  if (!range && !group) return false
  const ids = new Set(selectedRanges(rootEl).map((r) => r.blockId))
  if (range && ids.size === 0) ids.add(range.start.blockId)
  const paragraphs = [...ids].filter((id) => parseCellId(id) === null)
  const cellChanges = group ? cellsChangingKind(group.table, group.cells, kind) : []
  if (paragraphs.length === 0 && cellChanges.length === 0) return false
  pushHistory(caretPoint())
  for (const id of paragraphs) setContainerKindOp(doc.value, id, kind)
  if (group && cellChanges.length > 0) {
    setCellsKind(group.table, group.cells, kind)
    finishCellsOp(group)
    return true
  }
  refreshLayout({
    anchor: range?.start ?? null,
    anchorEnd: range?.end ?? null,
    force: true,
  })
  void nextTick(emitSelection)
  return true
}

function setBlockKind(kind: BlockKind): void {
  repeatable(() => applyBlockKind(kind))
}

function applyBold(): boolean {
  const rootEl = root.value
  if (!rootEl) return false
  const ranges = selectedRanges(rootEl)
  if (ranges.length === 0) return false
  const allBold = ranges.every((range) => {
    const container = findContainer(doc.value, range.blockId)
    if (!container) return false
    const p = prefixLength(range.blockId)
    return rangeIsBold(container, Math.max(0, range.from - p), Math.max(0, range.to - p))
  })
  forSelection((model, blockId, from, to) =>
    applyFormat(model, blockId, from, to, { bold: !allBold }),
  )
  return true
}

function toggleBold(): void {
  repeatable(applyBold)
}

function applyUnderline(): boolean {
  const rootEl = root.value
  if (!rootEl) return false
  const ranges = selectedRanges(rootEl)
  if (ranges.length === 0) return false
  const allUnderline = ranges.every((range) => {
    const container = findContainer(doc.value, range.blockId)
    if (!container) return false
    const p = prefixLength(range.blockId)
    return rangeIsUnderline(container, Math.max(0, range.from - p), Math.max(0, range.to - p))
  })
  forSelection((model, blockId, from, to) =>
    applyFormat(model, blockId, from, to, { underline: !allUnderline }),
  )
  return true
}

function toggleUnderline(): void {
  repeatable(applyUnderline)
}

function applyColor(hex: string | null): boolean {
  const rootEl = root.value
  if (!rootEl) return false
  if (selectedRanges(rootEl).length === 0) return false
  forSelection((model, blockId, from, to) =>
    applyFormat(model, blockId, from, to, {
      color: hex ? hex.replace(/^#/, '').toUpperCase() : null,
    }),
  )
  return true
}

function setColor(hex: string | null): void {
  repeatable(() => applyColor(hex))
}

/**
 * 接受 / 拒绝当前选区里的修订（没选中文字时按插入符所在的那一串算，见 revisionSpanAt）。
 *
 * 先问再改：一个修订都没覆盖到就一步都不动 —— 按钮虽然置灰，键盘路径与「点按钮到执行
 * 之间选区被改」这类竞态仍可能走到这里，记下一份什么都没变的撤销快照是最难查的那种错。
 *
 * 锚点按改之前的选区还回去：接受/拒绝插入修订时坐标不动，处理删除修订时后面会左移，
 * placeCaret 自己会找最近的落点。
 *
 * 返回「真的改了模型没有」—— 键盘那条路（`Ctrl+Alt+A` / `D`）拿它决定要不要弹提示。
 */
function resolveRevisionsOf(action: 'accept' | 'reject'): boolean {
  if (!props.editable) return false
  const rootEl = root.value
  if (!rootEl) return false

  /*
   * 要处理哪儿：**实时**选区优先 —— 有真选区就跨块全给上；只有插入符（折叠）时，
   * 按插入符所在的那一串修订算（Word 里点一下修订文字就能整段接受）。
   *
   * 这里**不能**像插入空格 / 加批注那样回退到 stickyRanges（「最近一次真选区」）：
   * 那两个入口的焦点会离开正文（下拉框、输入框），实时选区确实没了；而这两个按钮是
   * `@mousedown.prevent`，焦点一直在正文里，插入符就是最新的落点。回退过去反而会拿
   * 一份过期的选区当目标 —— 2026-09-15 实测踩到：选中「苏州」拒绝掉它之后，再点接受，
   * 用的还是那份旧选区（那儿已经没有修订了），于是这一下静默无效。
   */
  const targets: { blockId: string; from: number; to: number }[] = []
  const push = (blockId: string, from: number, to: number): void => {
    const p = prefixLength(blockId)
    const a = Math.max(0, from - p)
    targets.push({ blockId, from: a, to: Math.max(a, to - p) })
  }
  const live = selectedRanges(rootEl)
  if (live.length > 0) {
    for (const range of live) push(range.blockId, range.from, range.to)
  } else {
    const caret = selectionRange()?.start ?? lastCaret
    if (caret) push(caret.blockId, caret.offset, caret.offset)
  }

  // 先问再改：一处修订都没覆盖到就一步都不动（见上面的说明）
  const hits = targets.filter((target) => {
    const container = findContainer(doc.value, target.blockId)
    return container ? hasRevisions(container, target.from, target.to) : false
  })
  if (hits.length === 0) return false

  const before = selectionRange()
  pushHistory()
  for (const target of hits) {
    resolveRevisions(doc.value, target.blockId, target.from, target.to, action)
  }
  refreshLayout({
    anchor: before?.start ?? lastCaret,
    anchorEnd: before?.end ?? null,
    force: true,
  })
  // 修订没了，按钮得跟着变灰 —— 焦点没离开正文，不会自己来一次 selectionchange
  void nextTick(emitSelection)
  return true
}

/**
 * 把选中的数字改写成「千分位 + 两位小数」（alt+4）。
 * 不是合法数字、没有选中、或选中的不在同一段里时什么都不改，返回 false 交给调用方提示。
 */
function formatSelectionAsAmount(): boolean {
  if (!props.editable) return false
  const rootEl = root.value
  if (!rootEl) return false
  const ranges = selectedRanges(rootEl)
  const target = ranges.length === 1 ? ranges[0] : undefined
  if (!target) return false
  const container = findContainer(doc.value, target.blockId)
  if (!container) return false
  const p = prefixLength(target.blockId)
  const from = Math.max(0, target.from - p)
  const to = Math.max(from, target.to - p)
  if (to <= from) return false

  const amount = formatAmount(textOf(sliceStrict(container.inlines, from, to)))
  if (amount === null) return false

  // 新数字继承原文的格式（加粗／颜色／下划线／修订标记），与「改几个字」同一路数
  const first = sliceStrict(container.inlines, from, to).find(
    (inline): inline is TextInline => inline.t === 'text',
  )
  const piece: Inline = first ? { ...first, text: amount } : { t: 'text', text: amount }

  pushHistory()
  replaceRange(container, from, to, [piece])
  // 改完仍把这串数字选中：连着按几次结果稳定（12,345.60 再解析还是它自己）
  refreshLayout({
    anchor: { blockId: target.blockId, offset: p + from },
    anchorEnd: { blockId: target.blockId, offset: p + from + amount.length },
    force: true,
  })
  return true
}

const SPECIAL_SPACES: Record<'em' | 'en' | 'quarterEm', string> = {
  em: '\u2003',
  en: '\u2002',
  quarterEm: '\u2005',
}

/**
 * 在插入符处插入一个特殊空格；有选区时替换选区。
 * 插入符落在刚插入的空格之后（按块 id + 字符偏移找回，与其它结构性操作同一套）。
 */
function insertSpecialSpace(kind: 'em' | 'en' | 'quarterEm'): boolean {
  if (!props.editable) return false
  const rootEl = root.value
  if (!rootEl) return false

  /*
   * 落点的优先序：**实时选区 → 实时插入符 → 最近一次记录的选区 → 最近一次记录落点**。
   * 键盘路径（ctrl+alt+X/C/V）与「插入」页那三枚按钮（都是 @mousedown.prevent，焦点不离正文）
   * 永远落在前两项上；后两项只服务「焦点真的离开正文」的入口。
   *
   * 旧写法把 stickyRanges 排在插入符**之前**，于是「先在表格某格里选过字、再把插入符点到
   * 正文别处按这三枚键」时，空格全插进那个格子里（2026-09-16 用户实测，issues/20260916-1 第 3 条）。
   */
  const live = selectedRanges(rootEl)
  const caret = selectionRange()?.start ?? null
  const sticky = stickyRanges[0]
  const previous = lastCaret
  const target =
    live[0] !== undefined
      ? { blockId: live[0].blockId, from: live[0].from, to: live[0].to }
      : caret !== null
        ? { blockId: caret.blockId, from: caret.offset, to: caret.offset }
        : sticky !== undefined
          ? { blockId: sticky.blockId, from: sticky.from, to: sticky.to }
          : previous !== null
            ? { blockId: previous.blockId, from: previous.offset, to: previous.offset }
            : null
  if (target === null) return false
  const container = findContainer(doc.value, target.blockId)
  if (!container) return false

  const p = prefixLength(target.blockId)
  const from = Math.max(0, target.from - p)
  const to = Math.max(from, target.to - p)
  const char = SPECIAL_SPACES[kind]

  pushHistory(caret ?? previous)
  replaceRange(container, from, to, [{ t: 'text', text: char }])
  refreshLayout({
    anchor: { blockId: target.blockId, offset: p + from + char.length },
    force: true,
  })
  return true
}

/** 选中的文字加一条批注。焦点可能在批注输入框里，所以实时选区取不到就回退到记录 */
function addCommentOnSelection(text: string): number {
  const rootEl = root.value
  if (!rootEl) return -1
  const live = selectedRanges(rootEl)
  const ranges = live.length > 0 ? live : stickyRanges
  const target = ranges[0]
  if (!target) return -1
  const p = prefixLength(target.blockId)
  const id = addComment(
    doc.value,
    target.blockId,
    Math.max(0, target.from - p),
    Math.max(0, target.to - p),
    text,
    props.author,
    new Date().toISOString(),
  )
  if (id < 0) return -1
  activeCommentId.value = id
  // 批注已经落下，选区记录与那层「续命」高亮都该收掉了：
  // 留着会让下一条批注悄悄用上一次的选区。
  stickyRanges = []
  dropKeptSelection()
  refreshLayout({ force: true })
  return id
}

/** 给指定区间加批注（模型文字坐标） */
function addCommentAt(blockId: string, from: number, to: number, text: string): number {
  if (!findContainer(doc.value, blockId)) return -1
  const p = prefixLength(blockId)
  const id = addComment(
    doc.value,
    blockId,
    Math.max(0, from - p),
    Math.max(0, to - p),
    text,
    props.author,
    new Date().toISOString(),
  )
  if (id < 0) return -1
  activeCommentId.value = id
  refreshLayout({ force: true })
  return id
}

function replyComment(parentId: number, text: string): number {
  const id = replyCommentOp(doc.value, parentId, text, props.author, new Date().toISOString())
  refreshLayout({ force: true })
  return id
}

function removeComment(id: number): void {
  pushHistory()
  removeCommentOp(doc.value, id)
  if (activeCommentId.value === id) activeCommentId.value = null
  if (editingId.value === id) cancelEdit()
  refreshLayout({ force: true })
}

/** 改一条批注的内容。作者与时间不动 —— 改的是内容，不是谁在什么时候说的。 */
function editComment(id: number, text: string): boolean {
  const next = text.trim()
  if (next === '') return false
  pushHistory()
  if (!updateCommentOp(doc.value, id, next)) return false
  refreshLayout({ force: true })
  return true
}

/** 侧栏：把某条批注切成编辑态 / 退出编辑态 / 提交草稿 */
function startEdit(id: number, text: string): void {
  editingId.value = id
  editDraft.value = text
}

function cancelEdit(): void {
  editingId.value = null
  editDraft.value = ''
}

function commitEdit(id: number): void {
  if (editComment(id, editDraft.value)) cancelEdit()
}

/* -------------------------------------------------------------------------- */
/* 分页符 / 分节符                                                             */
/* -------------------------------------------------------------------------- */

/** 在落点所在块之后插入一个换页标记；落点取不到就追加到文末 */
function insertBreak(kind: BreakKind): void {
  if (!props.editable) return
  const point = caretPoint()
  pushHistory(point)
  insertBreakAfter(doc.value, point?.blockId, kind)
  refreshLayout({ anchor: point, force: true })
  // 新分节符会改「第 N 节 / 共 M 节」，补一次回显（点按钮已把焦点移出正文）
  void nextTick(emitSelection)
}

/** 点页间标记上的 × 时删掉那一枚分页符/分节符 */
function deleteBreak(blockId: string): void {
  if (!props.editable) return
  const point = caretPoint()
  pushHistory(point)
  if (!removeBreakOp(doc.value, blockId)) return
  refreshLayout({ anchor: point, force: true })
  void nextTick(emitSelection)
}

/** 工具栏按钮：在落点所在块之后插入分页符 / 分节符 */
function insertPageBreak(): void {
  insertBreak('page')
}

function insertSectionBreak(): void {
  insertBreak('section')
}

/* -------------------------------------------------------------------------- */
/* 节设置（W5）                                                                */
/* -------------------------------------------------------------------------- */

/**
 * 改光标所在节的设置。收尾三件与会话/节相关：pushHistory 记撤销点 → 重排
 * （版心高、页码显示都可能变） → 补一次 emitSelection（按钮/radio 会把焦点
 * 从正文拿走，selectionchange 未必再触发，不补工具条的选中态就停在旧值）。
 */
function updateSectionSetting(patch: SectionSettings): void {
  if (!props.editable) return
  const point = caretPoint()
  if (!point) return
  pushHistory(point)
  setSectionSetting(doc.value, sectionIndexOf(doc.value, point.blockId), patch)
  refreshLayout({ anchor: point, force: true })
  void nextTick(emitSelection)
}

function setSectionOrientation(orientation: PageOrientation): void {
  updateSectionSetting({ orientation })
}

function setSectionPageNumbers(on: boolean): void {
  updateSectionSetting({ pageNumbers: on })
}

function setSectionLinkPrevious(on: boolean): void {
  updateSectionSetting({ linkPrevious: on })
}

function setSectionRestartAtOne(on: boolean): void {
  updateSectionSetting({ restartAtOne: on })
}

/* -------------------------------------------------------------------------- */
/* 插入表格                                                                    */
/* -------------------------------------------------------------------------- */

/** 新表格的默认规格与上限：行数（不含 unit/note 行）、列数、最小行高（2 = 最小两行） */
const NEW_TABLE_ROWS = 2
const NEW_TABLE_COLUMNS = 3
const NEW_TABLE_MIN_LINES: 1 | 2 = 2
const MAX_TABLE_ROWS = 30
const MAX_TABLE_COLUMNS = 12

/**
 * 把界面上敲进来的行列数夹到合法区间。
 * 面板虽然写了 min/max，但数字是能直接敲进框里的（清空后还会给到 NaN），
 * 所以这里不信任输入：非法值回落默认值，越界值夹到边界。
 */
function clampCount(value: number, fallback: number, max: number): number {
  const n = Math.trunc(value)
  if (!Number.isFinite(n)) return fallback
  return Math.min(Math.max(1, n), max)
}

/**
 * 在当前块之后（找不到落点就追加到文末）插入一张空表，并把插入符放进第一个格子。
 *
 * 默认 2 行 × 3 列的空表、`minLines: 2`、`cantSplit: true`；行列数由调用方给
 * （demo 的「插入表格」面板让用户选）。增删行列与行高切换在 W4b。
 */
function insertTable(rows: number = NEW_TABLE_ROWS, columns: number = NEW_TABLE_COLUMNS): string {
  if (!props.editable) return ''
  const bodyRows = clampCount(rows, NEW_TABLE_ROWS, MAX_TABLE_ROWS)
  const cols = clampCount(columns, NEW_TABLE_COLUMNS, MAX_TABLE_COLUMNS)
  const point = caretPoint()
  pushHistory(point)

  // 落点在格子里时，插到那张表之后（不能插进格子里 —— 格子里放不下一个块）
  const anchorId = point?.blockId
  const cell = anchorId === undefined ? null : parseCellId(anchorId)
  const targetId = cell ? cell.tableId : anchorId
  const found =
    targetId === undefined ? -1 : doc.value.blocks.findIndex((block) => block.id === targetId)
  const at = found < 0 ? doc.value.blocks.length : found + 1

  const id = nextBlockId('tb')
  const table: TableBlock = {
    t: 'table',
    id,
    columns: cols,
    minLines: NEW_TABLE_MIN_LINES,
    cantSplit: true,
    rows: Array.from({ length: bodyRows }, () => ({
      role: 'body',
      cells: Array.from({ length: cols }, () => ({ inlines: [] })),
    })),
  }
  doc.value.blocks.splice(at, 0, table)

  // 插入符放进第一个格子的开头：格子 id + offset 0 就是它的锚点
  const first = cellId(id, 0, 0)
  refreshLayout({ anchor: { blockId: first, offset: 0 }, force: true })
  return id
}

/* -------------------------------------------------------------------------- */
/* 表格结构操作（W4b-1）                                                       */
/* -------------------------------------------------------------------------- */

/** 表格结构操作的落点。取不到（不可编辑 / 不在格子里）时为 null */
interface TableCaretTarget {
  tableId: string
  table: TableBlock
  /** 光标所在行在 rows 数组里的下标 */
  row: number
  /** 列下标；unit/note 行天然只有一格，恒为 0 */
  col: number
  /** 格内模型偏移（格子的前缀恒为 0，仍按现有写法扣一次更稳） */
  offset: number
  /** 改之前那一刻的落点，记进撤销栈用 */
  point: DisplayPoint
}

/** 取落点所在的表格与格子；不在格子里就直接判空（调用方据此空转、不 pushHistory） */
function tableTarget(): TableCaretTarget | null {
  if (!props.editable) return null
  const point = caretPoint()
  if (!point) return null
  const cell = parseCellId(point.blockId)
  if (!cell) return null
  const table = findTable(doc.value, cell.tableId)
  if (!table) return null
  return {
    tableId: cell.tableId,
    table,
    row: cell.row,
    col: cell.col,
    offset: Math.max(0, point.offset - prefixLength(point.blockId)),
    point,
  }
}

/** 某个 rows 下标之前有几个 body 行（即该行在 body 行里的序数） */
function bodyOrdinal(table: TableBlock, row: number): number {
  let count = 0
  for (let i = 0; i < row && i < table.rows.length; i += 1) {
    if (table.rows[i]?.role === 'body') count += 1
  }
  return count
}

/**
 * 表格结构操作的收尾：按**新下标**重算锚点 → 重排 → 把选区回显给工具条。
 *
 * `cellId` 里嵌的是行/列下标（`tb1.r2c1`），任何增删都会让其后的格子 id 整体位移，
 * 所以新锚点只能拿新下标重新组，不能沿用旧 id；偏移也要夹到该格的新长度以内。
 */
function finishTableOp(target: TableCaretTarget, row: number, col: number, offset: number): void {
  const table = target.table
  const r = Math.min(Math.max(0, row), Math.max(0, table.rows.length - 1))
  const cellCount = table.rows[r]?.cells.length ?? 0
  const c = Math.min(Math.max(0, col), Math.max(0, cellCount - 1))
  const cell = table.rows[r]?.cells[c]
  const at = Math.min(Math.max(0, offset), cell ? containerLength(cell) : 0)
  refreshLayout({
    anchor: { blockId: cellId(target.tableId, r, c), offset: at },
    force: true,
  })
  /*
   * 点工具条按钮 / radio 会把焦点从正文拿走，selectionchange 未必再触发 ——
   * 不补这一下，工具条的 radio 选中态与行列提示会停在旧值（用户看着就是「点了没反应」）。
   * emitSelection 只对外发事件、App 只把它存进 ref，不会有回环。
   */
  void nextTick(emitSelection)
}

/** 在光标所在行的上方 / 下方插入一行 body 行 */
function insertTableRow(where: 'above' | 'below'): void {
  const target = tableTarget()
  if (!target) return
  dropCellSelection()
  pushHistory(target.point)
  /*
   * 落点要夹进 body 区间：光标停在 unit 行时「上方插入行」的直觉落点是 0 号位，
   * 那会插到 unit 之前、破坏 unit → body… → note 的显示顺序（note 行同理）。
   */
  const at = bodyInsertIndex(target.table, where === 'above' ? target.row : target.row + 1)
  insertBodyRow(target.table, at)
  // 插在光标那一行之前 → 光标行整体 +1；插在它那一行（含）之后 → 光标行不动
  finishTableOp(target, at <= target.row ? target.row + 1 : target.row, target.col, target.offset)
}

/** 删除光标所在的 body 行；光标不在 body 行、或只剩一个 body 行时空转 */
function removeTableRow(): void {
  const target = tableTarget()
  if (!target) return
  // 这两种情形工具条上的按钮本来就是禁用态，这里再兜一道，权当防呆
  if (target.table.rows[target.row]?.role !== 'body') return
  if (bodyRowIndexes(target.table).length <= 1) return
  const ordinal = bodyOrdinal(target.table, target.row)
  dropCellSelection()
  pushHistory(target.point)
  removeBodyRow(target.table, target.row)
  // 删完后的 body 行序数取 min(k, bodyRows'-1)，再换算回 rows 数组下标
  const bodies = bodyRowIndexes(target.table)
  const row = bodies[Math.min(ordinal, bodies.length - 1)] ?? 0
  finishTableOp(target, row, target.col, target.offset)
}

/** 在光标所在列的左侧 / 右侧插入一列（只作用 body 行；unit/note 行整行一格，不动） */
function insertTableColumn(where: 'left' | 'right'): void {
  const target = tableTarget()
  if (!target) return
  dropCellSelection()
  pushHistory(target.point)
  insertColumn(target.table, where === 'left' ? target.col : target.col + 1)
  // 左侧插入：光标那一列变成 +1；右侧插入：光标那一列不动
  finishTableOp(target, target.row, where === 'left' ? target.col + 1 : target.col, target.offset)
}

/** 删除光标所在列；只剩一列时空转 */
function removeTableColumn(): void {
  const target = tableTarget()
  if (!target) return
  if (target.table.columns <= 1) return
  dropCellSelection()
  pushHistory(target.point)
  removeColumn(target.table, target.col)
  finishTableOp(target, target.row, Math.min(target.col, target.table.columns - 1), target.offset)
}

/** 行高两档（最小一行 / 最小两行；**只管正文行** —— 表头 / 附注行恒一行）；锚点不变 */
function applyTableMinLines(minLines: 1 | 2): boolean {
  const target = tableTarget()
  if (!target) return false
  if (target.table.minLines === minLines) return false
  pushHistory(target.point)
  setMinLines(target.table, minLines)
  finishTableOp(target, target.row, target.col, target.offset)
  return true
}

function setTableMinLines(minLines: 1 | 2): void {
  repeatable(() => applyTableMinLines(minLines))
}

/** radio 开关：增删表头行（unit，插在最前）/ 附注行（note，加在最后）；锚点按行号位移重算 */
function applyTableRoleRow(role: 'unit' | 'note', on: boolean): boolean {
  const target = tableTarget()
  if (!target) return false
  const table = target.table
  const index = table.rows.findIndex((row) => row.role === role)
  if (on && index >= 0) return false
  if (!on && index < 0) return false

  // 表头行插在最前、附注行加在最后，两者都会让行列下标整体位移 → 复选先收起（同增删行列）
  dropCellSelection()
  pushHistory(target.point)
  setRoleRow(table, role, on)
  if (role === 'unit') {
    if (on) {
      // 新行插在 0 号位，光标所在行整体 +1
      finishTableOp(target, target.row + 1, target.col, target.offset)
      return true
    }
    // 删掉表头行：光标就在它上面 → 落到第 0 行（此时是第一个 body 行）、偏移归 0；否则整体 -1
    if (target.row === index) finishTableOp(target, 0, target.col, 0)
    else finishTableOp(target, target.row - 1, target.col, target.offset)
    return true
  }

  // note 加在最后、删的也是最后一行，光标不在附注行时行号不变
  if (!on && target.row === index) {
    const bodies = bodyRowIndexes(table)
    finishTableOp(target, bodies[bodies.length - 1] ?? 0, target.col, target.offset)
    return true
  }
  finishTableOp(target, target.row, target.col, target.offset)
  return true
}

function setTableRoleRow(role: 'unit' | 'note', on: boolean): void {
  repeatable(() => applyTableRoleRow(role, on))
}

/* -------------------------------------------------------------------------- */
/* 表格：删除整表与两组对齐（W4b-2）                                            */
/* -------------------------------------------------------------------------- */

/**
 * 删除光标所在的整张表。**不二次确认**（与全仓库一致：不可逆操作靠 pushHistory 的撤销兜底）。
 * 锚点必须落到一个还存在的地方：优先被删块之前的第一个段落末尾，没有就取之后的第一个段落开头；
 * 两边都没有可落点的块时传 undefined —— 这条分支靠 refreshLayout 的兜底。
 */
function removeTable(): void {
  const target = tableTarget()
  if (!target) return
  const index = doc.value.blocks.findIndex((block) => block.id === target.tableId)
  if (index < 0) return
  // 整张表都没了，复选当然也留不住（refreshLayout 那侧还会兜一遍「表不在页面上」）
  dropCellSelection()
  pushHistory(target.point)
  if (!removeTableOp(doc.value, target.tableId)) return
  refreshLayout({ anchor: tableAnchorAfterRemoval(index), force: true })
  // 点按钮会把焦点从正文拿走，selectionchange 未必再触发；补这一下工具条才会随光标离开表格而收起
  void nextTick(emitSelection)
}

/** 删表后的落点：之前的第一个段落末尾 → 之后的第一个段落开头 → null（交给 refreshLayout 兜底） */
function tableAnchorAfterRemoval(index: number): DisplayPoint | null {
  for (let i = index - 1; i >= 0; i -= 1) {
    const block = doc.value.blocks[i]
    if (block?.t === 'textBlock') {
      return { blockId: block.id, offset: prefixLength(block.id) + blockLength(block) }
    }
  }
  for (let i = index; i < doc.value.blocks.length; i += 1) {
    const block = doc.value.blocks[i]
    if (block?.t === 'textBlock') return { blockId: block.id, offset: 0 }
  }
  return null
}

/**
 * 水平对齐三档（只有左 / 居中 / 右，不做两端对齐）；作用于**整格操作的目标格列表**。
 * 与当前**实际生效值**相同 → 清除该维覆盖（回默认：unit 右 / note 左 / body 跟该格样式）。
 */
function applyTableCellAlignH(value: Align): boolean {
  return applyCellsAlign('h', value)
}

/** 水平对齐三档的键盘 / 按钮入口。返回值 = 这一步真的改了模型没有（键盘那条路据此提示） */
function setTableCellAlignH(value: Align): boolean {
  return repeatable(() => applyTableCellAlignH(value))
}

/** 垂直对齐三档；同样「点当前值 = 回默认 top」，同样作用于目标格列表 */
function applyTableCellAlignV(value: CellVerticalAlign): boolean {
  return applyCellsAlign('v', value)
}

function setTableCellAlignV(value: CellVerticalAlign): boolean {
  return repeatable(() => applyTableCellAlignV(value))
}

/* -------------------------------------------------------------------------- */
/* 表格：整格复选（拖动刷选 + Ctrl+点击追加，W7）                                */
/* -------------------------------------------------------------------------- */

/**
 * 整格复选：一张表内的若干矩形格区间（模型坐标）。
 *
 * 它是**组件的交互状态**：既不进模型（模型只在批量操作执行时被改），也不进渲染函数 ——
 * `renderTableFragment` 与量测共用，让「选中」去改它产出的 HTML，量测缓存签名与逐块量测对账
 * 会一起崩。所以高亮是 **DOM 后处理**：组件按 `<td>` 上恒定的 `data-cell-id` 挂一个类名
 * （见 render/css.ts 的 CELL_SELECTION_CLASS），只改底色，不动任何几何。
 */
interface CellSelection {
  tableId: string
  /** 一个矩形一块；Ctrl+追加就是往里再加一块，「选中了哪些格」由它现推 */
  blocks: CellRect[]
}

/**
 * 正在按下的这一次指针手势。
 *
 * Word 的规则照抄在这里：**在格内按下拖动、只要没出这个格子就还是普通的选文字**，
 * 越过起点格的盒子才转成整格刷选（否则格内加粗、下划线这些就没法用了）。
 */
interface CellDrag {
  tableId: string
  /** 起点格在页面上的盒子（`<td>` 的矩形），用它判断「出格了没有」 */
  box: { left: number; top: number; right: number; bottom: number }
  /** 起点格（刷选时的第一个端点） */
  from: CellRef
  /** Ctrl 按下时：已有复选当基底；不按 Ctrl 时是空的（刷出来的矩形整块替换） */
  base: CellRect[]
  /** Ctrl 追加时的锚格（上一次整格操作的落点）；不按 Ctrl 时就是起点格 */
  anchor: CellRef
  ctrl: boolean
  /** 已经转成整格刷选（越过起点格至少一次） */
  brushing: boolean
}

/** 事件目标落在哪个格子里（不在格子里返回 null）；行列都归一化成模型坐标 */
function cellHit(
  target: EventTarget | null,
): { tableId: string; table: TableBlock; cell: CellRef } | null {
  if (!(target instanceof Element)) return null
  /*
   * 认格子的钩子是 `<td>` 上恒定的 `data-cell-id`，不是格内那层 div：指针落在格内的
   * **任何**地方都算这一格 —— 点在 div 之外的空白处（内边距、下端那半截空格子）时
   * `event.target` 就是 `<td>` 本身，只认 div 会把这些位置判成「不在格子里」。
   * （2026-09-16 实测踩到：格内文字只占上半截，瞄着格子中心按下时目标就是 `<td>`。）
   */
  const td = target.closest<HTMLElement>('td[data-cell-id]')
  if (!td) return null
  const parsed = parseCellId(td.dataset.cellId ?? '')
  if (!parsed) return null
  const table = findTable(doc.value, parsed.tableId)
  const row = table?.rows[parsed.row]
  if (!table || !row) return null
  const col = normalizeCellCol(table, parsed.row, parsed.col)
  // 渲染时凑矩形补出来的幻影格没有容器，落点选它没有意义（同 edit/table.ts 的 rowSlots）
  if (row.role === 'body' && !row.cells[col]) return null
  return { tableId: parsed.tableId, table, cell: { row: parsed.row, col } }
}

/**
 * 高亮重画：按模型坐标给命中的 `<td>` 挂类名。
 *
 * 只改 class、不碰几何、不重建 DOM —— 量测（与预览共用 renderTableFragment）与
 * 「逐块量测对账」都因此完全不受影响。同一张表跨页时高亮跟着各片段各自画（按坐标判，不按元素记）。
 */
function paintCellSelection(): void {
  const rootEl = root.value
  if (!rootEl) return
  const sel = cellSelection
  const hit = new Set<string>()
  if (sel) {
    const table = findTable(doc.value, sel.tableId)
    if (table) {
      for (const { row, col } of cellsInRects(table, sel.blocks)) {
        hit.add(cellId(sel.tableId, row, col))
      }
    }
  }
  for (const td of rootEl.querySelectorAll<HTMLElement>('td[data-cell-id]')) {
    td.classList.toggle(CELL_SELECTION_CLASS, hit.has(td.dataset.cellId ?? ''))
  }
}

/** 收起复选（只改状态与高亮，不 emit）。返回「原来有没有」 */
function dropCellSelection(): boolean {
  if (!cellSelection) return false
  cellSelection = null
  cellAnchor = null
  paintCellSelection()
  return true
}

/** 复选变了：重画高亮 + 把「选中了几格」报给调用方（不改 viewDoc / pages，也不重量测） */
function commitCellSelection(): void {
  paintCellSelection()
  emitSelection()
}

/**
 * 清空复选。退出条件（PLAN 13.2）：`Esc`、点正文别处、光标移出那张表、
 * 重排后那张表不在页面上了 —— 一律走到这里。
 */
function clearCellSelection(): void {
  if (dropCellSelection()) emitSelection()
}

/** 复选里**真实存在**的格（行优先、去重） */
function cellSelectionCells(): CellRef[] {
  const sel = cellSelection
  if (!sel) return []
  const table = findTable(doc.value, sel.tableId)
  return table ? cellsInRects(table, sel.blocks) : []
}

/**
 * 复选的概况（App 的「已选 N 格」与对齐/样式按钮的 active 态吃它）。
 *
 * 整批**一致**的值才带出去：不一致就省略那个字段 —— 调用方按它回显的话，
 * 选中的格里对不齐时按钮一个都不会亮（正是我们要的）。
 */
function cellSelectionSummary(): CellSelectionSummary | null {
  const sel = cellSelection
  if (!sel) return null
  const table = findTable(doc.value, sel.tableId)
  if (!table) return null
  const cells = cellSelectionCells()
  if (cells.length === 0) return null
  const spec = resolved.value
  const kinds = new Set<BlockKind>()
  const hs = new Set<Align>()
  const vs = new Set<CellVerticalAlign>()
  for (const { row, col } of cells) {
    const cell = table.rows[row]?.cells[col]
    const role = table.rows[row]?.role ?? 'body'
    const kind = cell ? storedCellKind(cell) : 'listItem'
    kinds.add(kind)
    hs.add(cell?.align?.h ?? defaultCellAlignH(role, spec.styles[kind].align))
    vs.add(cell?.align?.v ?? 'top')
  }
  const only = <T>(values: Set<T>): T | undefined =>
    values.size === 1 ? [...values][0] : undefined
  return {
    tableId: sel.tableId,
    count: cells.length,
    multiple: cells.length > 1,
    ...(only(kinds) ? { kind: only(kinds) } : {}),
    ...(only(hs) ? { alignH: only(hs) } : {}),
    ...(only(vs) ? { alignV: only(vs) } : {}),
  }
}

/**
 * 把原生选区折叠到它的起点。
 *
 * 「整格刷选」期间原生选区不再是选中态的载体（Word 也是整格高亮、不选文字），
 * 但它还会顺着指针一路扩 —— 所以转模式的那一刻收一次、拖完再收一次，
 * 保证刷选过后**没有残留的原生选区**（断言查的就是这一条）。
 */
function collapseNativeSelection(): void {
  const sel = document.getSelection()
  if (!sel || sel.rangeCount === 0) return
  const range = sel.getRangeAt(0)
  if (range.collapsed) return
  range.collapse(true)
  sel.removeAllRanges()
  sel.addRange(range)
}

function onPointerDown(event: PointerEvent): void {
  if (!props.editable || event.button !== 0) return
  cellDrag = null
  const hit = cellHit(event.target)
  if (!hit) {
    // 点正文别处：复选收起
    clearCellSelection()
    return
  }
  const td = (event.target as Element).closest('td') ?? (event.target as Element)
  const box = td.getBoundingClientRect()
  const ctrl = event.ctrlKey || event.metaKey
  const sameTable = cellSelection?.tableId === hit.tableId
  cellDrag = {
    tableId: hit.tableId,
    box: { left: box.left, top: box.top, right: box.right, bottom: box.bottom },
    from: hit.cell,
    base: ctrl && sameTable ? (cellSelection?.blocks ?? []).slice() : [],
    anchor: ctrl && sameTable && cellAnchor ? cellAnchor : hit.cell,
    ctrl,
    brushing: false,
  }
  // 不按 Ctrl：一按下就清（单击与「格内选文字」都算重新落点；Ctrl 那一下要拿旧复选当基底，不能清）
  if (!ctrl) clearCellSelection()
}

function onPointerMove(event: PointerEvent): void {
  const drag = cellDrag
  if (!drag) return
  if (!drag.brushing) {
    const inside =
      event.clientX >= drag.box.left &&
      event.clientX <= drag.box.right &&
      event.clientY >= drag.box.top &&
      event.clientY <= drag.box.bottom
    // 没出起点格：不接管，交给浏览器正常选文字（格内加粗、下划线全指着这条路）
    if (inside) return
    drag.brushing = true
    collapseNativeSelection()
  }
  // 刷选期间一直拦着原生选区扩展（拖到表格外也拦，否则选区会一路扩到正文里去）
  event.preventDefault()
  const hit = cellHit(event.target)
  if (!hit || hit.tableId !== drag.tableId) return
  cellSelection = { tableId: drag.tableId, blocks: [...drag.base, cellRectBetween(drag.anchor, hit.cell)] }
  cellAnchor = hit.cell
  commitCellSelection()
}

function onPointerUp(): void {
  const drag = cellDrag
  cellDrag = null
  if (!drag) return
  if (drag.brushing) {
    // 拖完再收一次：一路上被浏览器扩出来的原生选区在这里收干净
    collapseNativeSelection()
    commitCellSelection()
    return
  }
  if (!drag.ctrl) return
  /*
   * Ctrl+点击：点在**已选中的格**上 = 去掉包含它的那一块（Word 的行为）；
   * 否则以锚格为端点并上「锚格 ↔ 这一格」那块矩形。
   */
  const table = findTable(doc.value, drag.tableId)
  if (!table) return
  if (!cellSelection || cellSelection.tableId !== drag.tableId) {
    cellSelection = { tableId: drag.tableId, blocks: [cellRectBetween(drag.from, drag.from)] }
    cellAnchor = drag.from
  } else {
    const at = cellRectIndexOf(table, cellSelection.blocks, drag.from)
    if (at >= 0) {
      cellSelection.blocks.splice(at, 1)
      if (cellSelection.blocks.length === 0) {
        cellSelection = null
        cellAnchor = null
      }
    } else {
      cellSelection.blocks.push(cellRectBetween(drag.anchor, drag.from))
      cellAnchor = drag.from
    }
  }
  commitCellSelection()
}

function onPointerCancel(): void {
  cellDrag = null
}

/**
 * 整格操作的目标格列表（行优先、去重）。单选与多选走的是同一条代码：
 *   · 有整格复选 → 就是那一批（拖动刷选 / Ctrl+点击出来的）；
 *   · 没有复选 → 原生选区覆盖到的格子，一个都没有就回落到落点那一格。
 * 返回 null = 这一下不作用于任何格子（调用方据此走段落那条路或空转）。
 */
function cellTargets(): { tableId: string; table: TableBlock; cells: CellRef[] } | null {
  const sel = cellSelection
  if (sel) {
    const table = findTable(doc.value, sel.tableId)
    if (!table) return null
    const cells = cellsInRects(table, sel.blocks)
    return cells.length > 0 ? { tableId: sel.tableId, table, cells } : null
  }
  const rootEl = root.value
  const ids = new Set((rootEl ? selectedRanges(rootEl) : []).map((range) => range.blockId))
  const caret = caretPoint()
  if (caret) ids.add(caret.blockId)
  const cells: CellRef[] = []
  let tableId = ''
  let table: TableBlock | undefined
  for (const id of ids) {
    const parsed = parseCellId(id)
    if (!parsed) continue
    const found = findTable(doc.value, parsed.tableId)
    if (!found?.rows[parsed.row]) continue
    // 跨表选区只取第一张表：原生选区跨两张表没有先例，多做反而会在两条路之间来回抖
    if (tableId === '') {
      tableId = parsed.tableId
      table = found
    }
    if (parsed.tableId !== tableId) continue
    cells.push({ row: parsed.row, col: normalizeCellCol(found, parsed.row, parsed.col) })
  }
  if (!table || tableId === '') return null
  return { tableId, table, cells: sortCells(cells) }
}

/** 某一格该维的**实际生效值**（覆盖 ?? 角色 / 样式默认）—— 按钮回显与「再点同一个值就清除」都按它 */
function effectiveCellAlign(table: TableBlock, row: number, col: number, part: 'h' | 'v'): string {
  const cell = table.rows[row]?.cells[col]
  if (part === 'v') return cell?.align?.v ?? 'top'
  const role = table.rows[row]?.role ?? 'body'
  const kind = cell ? storedCellKind(cell) : 'listItem'
  return cell?.align?.h ?? defaultCellAlignH(role, resolved.value.styles[kind].align)
}

/**
 * 批量改一组格的某一维对齐。
 *
 * 规则与单选时同一条：整批的实际生效值都等于 value → 清除覆盖，否则一律写 value。
 * 一个格都不会变（本来就是这个值、又没有覆盖可清）→ 返回 false：一步都不许记撤销。
 */
function applyCellsAlign(part: 'h' | 'v', value: string): boolean {
  const group = cellTargets()
  if (!group) return false
  const { table, cells } = group
  const next = nextAlignValue(
    cells.map(({ row, col }) => effectiveCellAlign(table, row, col, part)),
    value,
  )
  if (cellsChangingAlign(table, cells, part, next).length === 0) return false
  pushHistory(caretPoint())
  setCellsAlign(table, cells, part, next)
  finishCellsOp(group)
  return true
}

/**
 * 批量格操作的收尾：锚点落到目标格之一 → 重排 → 补一次回显。
 *
 * 重排会重建 DOM，插入符必须按坐标放回去。点位优先沿用落点（复选把原生选区塌掉之后
 * 它仍在起点格里），落点不在这张表里就用第一格 —— 一次批量只 `pushHistory` 一次，
 * 撤销一步就能退回整批。
 */
function finishCellsOp(group: { tableId: string; table: TableBlock; cells: CellRef[] }): void {
  const first = group.cells[0]
  const caret = caretPoint()
  const parsed = caret ? parseCellId(caret.blockId) : null
  const inside = parsed?.tableId === group.tableId && group.table.rows[parsed.row] !== undefined
  const row = inside && parsed ? parsed.row : (first?.row ?? 0)
  const col = inside && parsed ? normalizeCellCol(group.table, parsed.row, parsed.col) : (first?.col ?? 0)
  const cell = group.table.rows[row]?.cells[col]
  const offset = caret && inside ? caret.offset - prefixLength(caret.blockId) : 0
  refreshLayout({
    anchor: {
      blockId: cellId(group.tableId, row, col),
      offset: Math.min(Math.max(0, offset), cell ? containerLength(cell) : 0),
    },
    force: true,
  })
  void nextTick(emitSelection)
}

/* -------------------------------------------------------------------------- */
/* 选区保持                                                                    */
/* -------------------------------------------------------------------------- */

/**
 * 点工具栏或批注输入框会把焦点从正文拿走，浏览器随即收起原生选区 —— 选中的底色
 * 就没了（批注其实照样加得上，因为区间记在 stickyRanges 里，丢的只是视觉）。
 *
 * 这里用 CSS Custom Highlight API 单独画一层高亮：它不动 DOM，也就不必重排、
 * 不会碰到正在编辑的内容。不支持的浏览器直接跳过，降级成「没有高亮」。
 */
type HighlightRegistry = Map<string, Highlight>

function highlightRegistry(): HighlightRegistry | null {
  if (typeof CSS === 'undefined') return null
  const registry = (CSS as unknown as { highlights?: HighlightRegistry }).highlights
  return registry ?? null
}

function keepSelection(): void {
  if (!props.editable) return
  const registry = highlightRegistry()
  if (!registry) return
  const sel = document.getSelection()
  if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return
  const range = sel.getRangeAt(0)
  if (!fragmentOf(range.startContainer)) return
  registry.set(KEEP_SELECTION_HIGHLIGHT, new Highlight(range.cloneRange()))
}

function dropKeptSelection(): void {
  highlightRegistry()?.delete(KEEP_SELECTION_HIGHLIGHT)
}

function undo(): void {
  const entry = undoStack.pop()
  if (!entry) return
  redoStack.push({ doc: cloneDoc(doc.value), caret: lastCaret, end: null })
  doc.value = entry.doc
  refreshLayout({ anchor: entry.caret, anchorEnd: entry.end, force: true })
}

function redo(): void {
  const entry = redoStack.pop()
  if (!entry) return
  undoStack.push({ doc: cloneDoc(doc.value), caret: lastCaret, end: null })
  doc.value = entry.doc
  refreshLayout({ anchor: entry.caret, anchorEnd: entry.end, force: true })
}

/** 点侧栏 → 高亮正文锚点并滚动到它 */
async function focusComment(id: number): Promise<void> {
  activeCommentId.value = activeCommentId.value === id ? null : id
  await nextTick()
  root.value
    ?.querySelector(`.wtp-comment[data-comment="${id}"]`)
    ?.scrollIntoView({ block: 'center', behavior: 'smooth' })
}

/* -------------------------------------------------------------------------- */
/* 导航窗格：大纲                                                              */
/* -------------------------------------------------------------------------- */

/**
 * 大纲跑在渲染快照（viewDoc / viewNumbering）上，不是编辑中的模型：
 * 打字时快照不动，左栏就不会跟着重渲染。只有指纹变了才对外发一次。
 */
function syncOutline(): void {
  const entries = buildOutline(viewDoc.value, viewNumbering.value)
  const sig = outlineSignature(entries)
  if (sig === outlineSig) return
  outlineSig = sig
  emit('outline-change', entries)
}

/** 点导航条目：把该块滚到可视区中间；可编辑时再把插入符放到自动编号之后 */
function focusBlock(blockId: string): void {
  const rootEl = root.value
  if (!rootEl) return
  const frag = rootEl.querySelector<HTMLElement>(`[data-block-id="${CSS.escape(blockId)}"]`)
  if (!frag) return
  if (props.editable) {
    placeCaret(rootEl, {
      blockId,
      offset: Number(frag.dataset.from ?? '0') + prefixLengthOf(frag),
    })
  }
  frag.scrollIntoView({ block: 'center' })
}

/* -------------------------------------------------------------------------- */
/* 查找与替换                                                                  */
/* -------------------------------------------------------------------------- */

/** 当前选区覆盖到的每一块（模型文字坐标）；取不到就是空数组 */
function getSelectionScope(): SearchScope[] {
  const rootEl = root.value
  if (!rootEl) return []
  const ranges = selectedRanges(rootEl)
  if (ranges.length === 0) return []
  const numbering = numberingOf(doc.value)
  return ranges.map((range) => {
    const prefix = (numbering.get(range.blockId) ?? '').length
    const from = Math.max(0, range.from - prefix)
    return { blockId: range.blockId, from, to: Math.max(from, range.to - prefix) }
  })
}

/**
 * 一个匹配在页面上对应的 Range。一个匹配可能横跨两个分页片段（同一块被切开），
 * 这时必须**按片段切成多个子 Range** —— 用「起点在 A 片、终点在 B 片」的单个 Range
 * 会把两页之间的换页标记与页脚一并圈进去。
 */
function matchRanges(match: Match, numbering: Map<string, string>): Range[] {
  const rootEl = root.value
  if (!rootEl) return []
  const prefix = (numbering.get(match.blockId) ?? '').length
  const lo = prefix + match.from
  const hi = prefix + match.to
  if (hi <= lo) return []

  const out: Range[] = []
  const frags = rootEl.querySelectorAll<HTMLElement>(
    `[data-block-id="${CSS.escape(match.blockId)}"]`,
  )
  for (const frag of frags) {
    const from = Number(frag.dataset.from ?? '0')
    const raw = frag.dataset.to
    const to = raw !== undefined && raw !== '' ? Number(raw) : from + (frag.textContent ?? '').length
    const a = Math.max(lo, from)
    const b = Math.min(hi, to)
    if (b <= a) continue
    const start = offsetToPoint(frag, a - from)
    const end = offsetToPoint(frag, b - from)
    if (!start || !end) continue
    const range = document.createRange()
    range.setStart(start.node, start.offset)
    range.setEnd(end.node, end.offset)
    out.push(range)
  }
  return out
}

/**
 * 重画高亮。浏览器不支持 Custom Highlight API 时降级成「没有高亮」，匹配与替换照常。
 * 只在「DOM 代次 / 匹配集合 / 当前匹配」三样里有变化时才重画：Range 挂在具体节点上，
 * 页面重建过就必须重画，其它时候同样的匹配重画一遍纯属浪费。
 */
function paintSearchHighlights(): void {
  const registry = highlightRegistry()
  if (!registry) return
  const key = `${domGen}|${searchIndex}|${searchMatches
    .map((m) => `${m.blockId}:${m.from}:${m.to}`)
    .join(',')}`
  if (key === searchPaintKey) return
  searchPaintKey = key
  if (searchMatches.length === 0) {
    registry.delete(SEARCH_HIGHLIGHT)
    registry.delete(SEARCH_CURRENT_HIGHLIGHT)
    return
  }
  const numbering = numberingOf(doc.value)
  const normal: Range[] = []
  const current: Range[] = []
  searchMatches.forEach((match, index) => {
    const ranges = matchRanges(match, numbering)
    if (ranges.length === 0) return
    if (index === searchIndex) current.push(...ranges)
    else normal.push(...ranges)
  })
  if (normal.length > 0) registry.set(SEARCH_HIGHLIGHT, new Highlight(...normal))
  else registry.delete(SEARCH_HIGHLIGHT)
  if (current.length > 0) registry.set(SEARCH_CURRENT_HIGHLIGHT, new Highlight(...current))
  else registry.delete(SEARCH_CURRENT_HIGHLIGHT)
}

function emitSearchState(): void {
  emit('search-state', {
    total: searchMatches.length,
    current: searchIndex >= 0 ? searchIndex + 1 : 0,
    error: searchError,
  })
}

/**
 * 重跑匹配并重画。reset 为真表示查询本身变了（回到第一处）；
 * 为假表示只是文档在改（尽量留在原来那一处，免得每敲一个字高亮都跳回文首）。
 */
function recomputeSearch(reset: boolean): void {
  searchError = validateQuery(searchQuery, searchRegex)
  searchMatches =
    searchError !== null || searchQuery === ''
      ? []
      : findMatches(doc.value, searchQuery, {
          regex: searchRegex,
          ...(searchScope !== undefined ? { scope: searchScope } : {}),
        })
  if (searchMatches.length === 0) searchIndex = -1
  else if (reset || searchIndex < 0) searchIndex = 0
  else searchIndex = Math.min(searchIndex, searchMatches.length - 1)
  // 高亮的 Range 要挂在新 DOM 上：重排后的节点是新建的，得等 Vue 渲染完再画
  void nextTick(paintSearchHighlights)
  emitSearchState()
}

/** 把当前匹配滚到可视区中间（滚动目标取它的第一片） */
function scrollToCurrentMatch(): void {
  const match = searchMatches[searchIndex]
  if (!match) return
  const range = matchRanges(match, numberingOf(doc.value))[0]
  if (!range) return
  const node = range.startContainer
  const el = node.nodeType === 1 ? (node as HTMLElement) : node.parentElement
  el?.scrollIntoView({ block: 'center' })
}

function setSearch(query: string, opts: SearchOptions = {}): void {
  searchQuery = query
  searchRegex = opts.regex === true
  searchScope = opts.scope
  searchOn = true
  recomputeSearch(true)
}

function moveMatch(step: number): void {
  const total = searchMatches.length
  if (total === 0) return
  searchIndex = (searchIndex + step + total) % total
  void nextTick(() => {
    paintSearchHighlights()
    scrollToCurrentMatch()
  })
  emitSearchState()
}

function nextMatch(): void {
  moveMatch(1)
}

function prevMatch(): void {
  moveMatch(-1)
}

/** 替换用的修订标记：kind 由 replaceMatches 按用途覆盖，这里只负责唯一 id 与作者/时间 */
const replaceRev = (): RevMark => mark('ins')

/**
 * 替换之后插入符该落在哪里：新文字的末尾（模型坐标换算成显示坐标，加自动编号前缀）。
 * 修订模式下新文字落在被删文字之后（见 replaceMatches），所以从 to 起算；否则原地替换，从 from 起算。
 */
function caretAfterReplace(match: Match, text: string): DisplayPoint {
  const at = (props.trackChanges ? match.to : match.from) + text.length
  return { blockId: match.blockId, offset: prefixLength(match.blockId) + at }
}

function replaceCurrent(text: string): void {
  if (!props.editable) return
  const match = searchMatches[searchIndex]
  if (!match) return
  pushHistory()
  replaceMatches(doc.value, [match], text, props.trackChanges ? replaceRev : undefined)
  // 必须给 anchor：force 重排会重建 DOM，不给锚点插入符就丢，替换完接着敲字会插到段首
  refreshLayout({ force: true, anchor: caretAfterReplace(match, text) })
}

function replaceAll(text: string): void {
  if (!props.editable || searchMatches.length === 0) return
  // 停在第一处（与下面 searchIndex 归零一致）；首处的区间不会被后面的替换挪动
  const first = searchMatches[0]
  pushHistory()
  replaceMatches(doc.value, searchMatches, text, props.trackChanges ? replaceRev : undefined)
  searchIndex = 0
  refreshLayout({ force: true, anchor: first ? caretAfterReplace(first, text) : null })
}

function clearSearch(): void {
  searchOn = false
  searchQuery = ''
  searchScope = undefined
  searchMatches = []
  searchIndex = -1
  searchError = null
  const registry = highlightRegistry()
  registry?.delete(SEARCH_HIGHLIGHT)
  registry?.delete(SEARCH_CURRENT_HIGHLIGHT)
  // 高亮已经手动撤掉了，签名也要作废 —— 否则下一轮画出同样的匹配会被跳过
  searchPaintKey = ''
  emit('search-state', { total: 0, current: 0, error: null })
}

/* -------------------------------------------------------------------------- */
/* 编辑器开关（模型里的 ::editor）                                              */
/* -------------------------------------------------------------------------- */

/**
 * 把模型里的编辑器开关报给调用方。**只在模型载入 / 重建之后发**（挂载与 source/model 变更），
 * 让「模型赢」：调用方用自己的默认状态顶不掉模型里写的值。
 */
function emitEditorFlags(): void {
  emit('editor-flags', resolveEditorFlags(doc.value.editor))
}

/**
 * 把开关写回模型（供调用方在用户改开关时调）。缺省值要**删字段** —— 默认值不落字段是
 * 「模型 → md → 模型」字节稳定的前提；全默认时整个 editor 字段都不留。
 *
 * 这里**不发 editor-flags**：调用方是「先改自己的状态、再回写模型」，再发事件就形成回环，
 * 把用户刚改的值覆盖回旧的。
 */
function setEditorFlags(flags: EditorFlags): void {
  const next: EditorSettings = {}
  if (flags.trackChanges) next.trackChanges = true
  if (!flags.nav) next.nav = false
  if (next.trackChanges === undefined && next.nav === undefined) delete doc.value.editor
  else doc.value.editor = next
}

/* -------------------------------------------------------------------------- */
/* 生命周期                                                                    */
/* -------------------------------------------------------------------------- */

onMounted(async () => {
  // 必须等字体就绪再量。仿宋/华文中宋是本机字体，Chromium 仍会异步加载：
  // 若在就绪前测量，量到的是回退字体的行盒，分页会整体偏掉且不会自愈。
  if (typeof document !== 'undefined' && 'fonts' in document) {
    try {
      await document.fonts.ready
    } catch {
      // 字体加载失败就按回退度量，不阻断渲染
    }
  }
  document.addEventListener('selectionchange', onSelectionChange)
  /*
   * 整格刷选挂在 document 上：指针一旦拖出起点格，后续的移动与松手可能落在版心之外
   * （别的页、工具栏、窗口外），只挂在 .wtp-content 上会丢事件。pointerdown 例外 ——
   * 它只挂版心（见模板），点工具栏/侧栏时复选不该被清掉。
   */
  document.addEventListener('pointermove', onPointerMove)
  document.addEventListener('pointerup', onPointerUp)
  document.addEventListener('pointercancel', onPointerCancel)
  refreshLayout({ force: true })
  emitEditorFlags()
})

onBeforeUnmount(() => {
  document.removeEventListener('selectionchange', onSelectionChange)
  document.removeEventListener('pointermove', onPointerMove)
  document.removeEventListener('pointerup', onPointerUp)
  document.removeEventListener('pointercancel', onPointerCancel)
  if (compSyncTimer !== null) clearTimeout(compSyncTimer)
})

watch([() => props.source, () => props.model], () => {
  doc.value = props.model ?? parseMd(props.source, { author: props.author })
  clearMeasureCache(cache)
  undoStack.length = 0
  redoStack.length = 0
  refreshLayout({ force: true })
  // 模型换了，开关跟着换（模型是权威）
  emitEditorFlags()
})

watch(
  () => props.spec,
  () => {
    /*
     * 换规格表（切文件模板）必须清量测缓存并整篇重量：页边距一变版心就变，
     * measure.ts 里所有行数与行高全部失效 —— 不清缓存就会拿旧版心的量测值算页码，
     * 而新版心越接近旧版心，这个错越不容易被看出来。
     *
     * 落点要在重量之前取：整篇的行边界都会挪，插入符与选区只能靠「块 id + 字符偏移」
     * 还回去（DOM 节点会被重建），不能靠节点引用。
     */
    const range = selectionRange() ?? rangeOfSticky()
    clearMeasureCache(cache)
    refreshLayout({
      anchor: range?.start ?? lastCaret,
      anchorEnd: range?.end ?? null,
      force: true,
    })
  },
  { deep: true },
)

watch(
  () => props.editable,
  (on) => {
    // 只读预览上做替换没有意义，而且切模式时 props.source 会把模型整个重新解析一遍，
    // 旧高亮对应的节点早没了 —— 清掉最干净。
    if (!on) clearSearch()
  },
)

async function exportDocx(): Promise<Blob> {
  return toBlob(doc.value, resolved.value, { title: '公文' })
}

async function downloadDocx(filename = '公文.docx'): Promise<void> {
  const blob = await exportDocx()
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  link.click()
  URL.revokeObjectURL(url)
}

defineExpose({
  repaginate,
  exportDocx,
  downloadDocx,
  /** 当前模型，供调用方做进一步处理 */
  getModel: (): DocModel => doc.value,
  /** 当前生效的完整规格表 */
  getSpec: (): Spec => resolved.value,
  /** 最近一次分页用到的量测值（行数、行高、段距），排错用 */
  getMeasurements: (): MeasuredItem[] => measured.value,
  pageCount: (): number => pages.value.length,
  // 编辑器开关（模型里的 ::editor），调用方在用户改开关时回写
  setEditorFlags,
  // 编辑层
  setBlockKind,
  toggleBold,
  toggleUnderline,
  setColor,
  // 接受 / 拒绝修订
  resolveRevisions: resolveRevisionsOf,
  formatSelectionAsAmount,
  insertSpecialSpace,
  addCommentOnSelection,
  addCommentAt,
  replyComment,
  removeComment,
  editComment,
  insertPageBreak,
  insertSectionBreak,
  // 节设置（W5）
  setSectionOrientation,
  setSectionPageNumbers,
  setSectionLinkPrevious,
  setSectionRestartAtOne,
  insertTable,
  // 表格结构操作（W4b-1）
  insertTableRow,
  removeTableRow,
  insertTableColumn,
  removeTableColumn,
  setTableMinLines,
  setTableRoleRow,
  // 表格：删除整表与两组对齐（W4b-2）
  removeTable,
  setTableCellAlignH,
  setTableCellAlignV,
  /** 整格复选里**真实存在**的格（行优先、模型坐标）—— 只读镜像，供验收脚本核对（W7） */
  getCellSelection: (): { tableId: string; cells: CellRef[] } | null => {
    const sel = cellSelection
    if (!sel) return null
    return { tableId: sel.tableId, cells: cellSelectionCells() }
  },
  clearCellSelection,
  deleteBreak,
  keepSelection,
  dropKeptSelection,
  undo,
  redo,
  canUndo: (): boolean => undoStack.length > 0,
  canRedo: (): boolean => redoStack.length > 0,
  focusComment,
  // 查找与替换（面板由调用方画）
  getSelectionScope,
  setSearch,
  nextMatch,
  prevMatch,
  replaceCurrent,
  replaceAll,
  clearSearch,
  // 导航窗格
  focusBlock,
})
</script>

<template>
  <div ref="root" class="wtp-root">
    <div class="wtp-pages">
      <template v-for="page in pages" :key="`${page.sectionIndex}-${page.pageNumber}`">
        <div class="wtp-page" :style="pageStyle(page)">
          <div
            class="wtp-content"
            :contenteditable="editable ? 'true' : undefined"
            :spellcheck="editable ? 'false' : undefined"
            @pointerdown="onPointerDown"
            @input="onInput"
            @keydown="onKeydown"
            @beforeinput="onBeforeInput"
            @paste="onPaste"
            @compositionstart="onCompositionStart"
            @compositionend="onCompositionEnd"
          >
            <div
              v-for="(frag, index) in page.fragments"
              :key="fragmentKey(frag)"
              :class="fragmentClass(frag)"
              :style="fragmentStyle(frag, index === 0)"
              :data-block-id="frag.rowFrom === undefined ? frag.blockId : undefined"
              :data-from="frag.rowFrom === undefined ? frag.from : undefined"
              :data-to="frag.rowFrom === undefined ? frag.to : undefined"
              :data-table-id="frag.rowFrom === undefined ? undefined : frag.blockId"
              :data-row-from="frag.rowFrom"
              :data-row-to="frag.rowTo"
              :data-continuation="frag.continuation ? '1' : undefined"
              v-html="fragmentHtml(frag)"
            />
          </div>
          <!-- wtp-footer 提供排版（= 内置「页脚」样式），wtp-page-number 只负责定位 -->
          <div
            v-if="page.showPageNumber"
            class="wtp-page-number wtp-footer"
            :style="pageNumberStyle(page)"
          >
            {{ page.pageNumber }}
          </div>
        </div>
        <!--
          换页标记，可以有不止一枚（分页符 + 分节符连在一起时两枚都落在同一页底部）。
          放在两页之间的空隙里而不是版心内：既不挤占版心高度（分页算术不必为它记账），
          也不会盖住正文。有它才能看见「这里插了一个分节符」。
        -->
        <div
          v-for="mark in page.breaks"
          :key="mark.blockId"
          class="wtp-break"
          :data-break-id="mark.blockId"
        >
          <span class="wtp-break-label">{{
            mark.kind === 'section' ? '分节符（下一页）' : '分页符'
          }}</span>
          <button
            v-if="editable"
            type="button"
            class="wtp-break-del"
            title="删除这个换页标记"
            @click="deleteBreak(mark.blockId)"
          >
            ×
          </button>
        </div>
      </template>
      <div v-if="pages.length === 0" class="wtp-page" :style="pageStyle(EMPTY_PAGE)">
        <div class="wtp-content" />
        <div class="wtp-page-number wtp-footer" :style="pageNumberStyle(EMPTY_PAGE)">1</div>
      </div>
    </div>

    <!-- 审阅侧栏：文档里没有批注时整块不占位 -->
    <aside v-if="comments.length > 0" class="wtp-comments">
      <div class="wtp-comments-title">批注 {{ comments.length }}</div>
      <ul>
        <li v-for="c in comments" :key="c.id" class="wtp-comment-row" :data-comment-id="c.id">
          <button
            type="button"
            class="wtp-comment-item"
            :class="{ 'is-active': c.id === activeCommentId }"
            @click="focusComment(c.id)"
          >
            <span class="wtp-comment-meta">
              #{{ c.id + 1 }} {{ c.author
              }}<template v-if="formatDate(c.date)"> · {{ formatDate(c.date) }}</template>
              <em v-if="c.resolved">已解决</em>
            </span>
            <span class="wtp-comment-scope">「{{ c.scope }}」</span>
            <span v-if="editingId !== c.id" class="wtp-comment-text">{{ c.text }}</span>
          </button>
          <div v-if="editingId === c.id" class="wtp-comment-edit">
            <input
              v-model="editDraft"
              type="text"
              spellcheck="false"
              @keydown.enter.prevent="commitEdit(c.id)"
              @keydown.esc.prevent="cancelEdit"
            />
            <button type="button" class="primary" @click="commitEdit(c.id)">保存</button>
            <button type="button" @click="cancelEdit">取消</button>
          </div>
          <div class="wtp-comment-actions">
            <button type="button" @click="startEdit(c.id, c.text)">编辑</button>
            <button type="button" @click="removeComment(c.id)">删除</button>
          </div>
          <ul v-if="c.replies.length > 0" class="wtp-replies">
            <li v-for="r in c.replies" :key="r.id">
              <span class="wtp-comment-meta"
                >{{ r.author
                }}<template v-if="formatDate(r.date)"> · {{ formatDate(r.date) }}</template></span
              >
              <span class="wtp-comment-text">{{ r.text }}</span>
            </li>
          </ul>
        </li>
      </ul>
    </aside>

    <!-- 测量容器挂载点：必须真实参与布局，否则量不到行盒 -->
    <div ref="host" class="wtp-measure-root" aria-hidden="true" />
  </div>
</template>

<style scoped>
.wtp-root {
  display: flex;
  align-items: flex-start;
  gap: 16px;
}
.wtp-pages {
  flex: 1;
  min-width: 0;
}
.wtp-measure-root {
  position: absolute;
  width: 0;
  height: 0;
  overflow: hidden;
  visibility: hidden;
}

.wtp-comments {
  position: sticky;
  top: 0;
  flex: 0 0 264px;
  max-height: 100%;
  overflow: auto;
  box-sizing: border-box;
  border: 1px solid #e4e6ea;
  border-radius: 6px;
  background: #fff;
  font-size: 12px;
  color: #4a4f56;
}
.wtp-comments-title {
  padding: 8px 10px;
  border-bottom: 1px solid #e4e6ea;
  background: #f6f7f9;
  font-weight: 600;
}
.wtp-comments ul {
  margin: 0;
  padding: 0;
  list-style: none;
}
.wtp-comments li + li {
  border-top: 1px solid #eef0f3;
}
/* 只有「正文那条」是整块可点的条目按钮；编辑/删除是下面独立的一排小按钮 */
.wtp-comments .wtp-comment-item {
  display: block;
  width: 100%;
  padding: 8px 10px;
  border: 0;
  background: transparent;
  font: inherit;
  color: inherit;
  text-align: left;
  cursor: pointer;
}
.wtp-comments .wtp-comment-item:hover {
  background: #f8f9fb;
}
.wtp-comments .wtp-comment-item.is-active {
  background: rgba(255, 213, 0, 0.35);
}
.wtp-comment-edit {
  display: flex;
  gap: 4px;
  padding: 0 10px 6px;
}
.wtp-comment-edit input {
  flex: 1;
  min-width: 0;
  padding: 3px 6px;
  border: 1px solid #c8ccd2;
  border-radius: 4px;
  font: inherit;
}
.wtp-comment-edit button,
.wtp-comment-actions button {
  padding: 3px 8px;
  border: 1px solid #c8ccd2;
  border-radius: 4px;
  background: #fff;
  color: #4a4f56;
  font: inherit;
  cursor: pointer;
}
.wtp-comment-edit button.primary {
  border-color: #1f6feb;
  background: #1f6feb;
  color: #fff;
}
.wtp-comment-actions {
  display: flex;
  gap: 6px;
  padding: 0 10px 8px;
}
.wtp-comment-actions button:hover {
  border-color: #8a9099;
}
/* 页间换页标记：放在两页之间的空隙里，不占版心。
   flex-basis 100% 是必须的：页带是可换行的横排（见 lib/render/css.ts），
   标记若不占满整行，并排时它会被塞进两页之间，把第二页顶到下一行去。
   负外边距照旧 —— 它负责把这个不占版心的标记压进页间那道 18px 的缝里。 */
.wtp-break {
  display: flex;
  align-items: center;
  gap: 6px;
  flex: 0 0 100%;
  margin: -6px 0;
  color: #8a9099;
  font-size: 11px;
  line-height: 1;
}
.wtp-break-label {
  padding: 1px 8px;
  border: 1px dashed #b6bcc4;
  border-radius: 999px;
  background: #f3f4f6;
  white-space: nowrap;
}
.wtp-break-del {
  width: 18px;
  height: 18px;
  padding: 0;
  border: 1px solid #c8ccd2;
  border-radius: 999px;
  background: #fff;
  color: #8a9099;
  line-height: 1;
  cursor: pointer;
}
.wtp-break-del:hover {
  border-color: #e0a9a4;
  color: #b3261e;
}
.wtp-replies {
  padding: 0 10px 8px 22px !important;
}
.wtp-replies li {
  border: 0 !important;
}
.wtp-comment-meta {
  display: block;
  color: #8a9099;
  font-size: 11px;
}
.wtp-comment-meta em {
  margin-left: 6px;
  padding: 0 4px;
  border-radius: 3px;
  background: #e8f2e9;
  color: #2f7a3f;
  font-style: normal;
}
.wtp-comment-scope {
  display: block;
  margin-top: 2px;
  color: #7a6019;
}
.wtp-comment-text {
  display: block;
  margin-top: 2px;
  line-height: 1.6;
}
</style>
