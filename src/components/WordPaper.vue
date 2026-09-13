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
  addComment,
  applyFormat,
  cloneDoc,
  findBlock,
  insertBreakAfter,
  insertText,
  mergeIntoPrevious,
  rangeColor,
  rangeIsBold,
  rangeIsUnderline,
  removeBreak as removeBreakOp,
  removeComment as removeCommentOp,
  replyComment as replyCommentOp,
  replaceRange,
  setBlockKind as setBlockKindOp,
  sliceStrict,
  splitBlock,
  updateComment as updateCommentOp,
} from '../lib/edit/model'
import type { EditorSelection } from '../lib/edit/model'
import { parseMd } from '../lib/md/parse'
import { computeNumbering } from '../lib/numbering'
import {
  KEEP_SELECTION_HIGHLIGHT,
  SEARCH_CURRENT_HIGHLIGHT,
  SEARCH_HIGHLIGHT,
  injectCss,
} from '../lib/render/css'
import { renderInlinesHtml } from '../lib/render/html'
import { clearMeasureCache, measureDocument } from '../lib/render/measure'
import type { MeasureCache } from '../lib/render/measure'
import { paginate } from '../lib/render/paginate'
import type { BreakKind, MeasuredItem, PageFragment, PageLayout } from '../lib/render/paginate'
import { contentBoxPx, resolveSpec } from '../lib/spec'
import type { BlockKind, DeepPartial, Spec } from '../lib/spec'
import { commentScopes } from '../lib/types'
import type { DocModel, Inline, RevMark, TextBlock, TextInline } from '../lib/types'

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
  }>(),
  { source: '', author: '管理员', editable: false, trackChanges: false },
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

const host = ref<HTMLElement | null>(null)
const root = ref<HTMLElement | null>(null)
const pages = shallowRef<PageLayout[]>([])
/** 最近一次分页实际用到的量测值，排错时用来和渲染结果对账 */
const measured = shallowRef<MeasuredItem[]>([])
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

/** 与预览分页同源的切片（保留跨片批注锚点，让高亮画全） */
function sliceForView(inlines: readonly Inline[], from: number, to: number): Inline[] {
  const out: Inline[] = []
  let cursor = 0
  for (const inline of inlines) {
    if (inline.t !== 'text') {
      out.push(inline)
      continue
    }
    const start = cursor
    const end = cursor + inline.text.length
    cursor = end
    if (end <= from || start >= to) continue
    out.push({
      ...inline,
      text: inline.text.slice(Math.max(from, start) - start, Math.min(to, end) - start),
    })
  }
  return out
}

/**
 * 渲染一个分页片段。
 * 片段偏移用的是「显示文字」坐标系（含自动编号前缀），模型里不含前缀，
 * 所以要先把前缀长度扣掉；只有首片才带前缀。
 */
function fragmentHtml(frag: PageFragment): string {
  const block = blocksById.value.get(frag.blockId)
  if (!block) return ''
  const prefix = viewNumbering.value.get(frag.blockId) ?? ''
  const p = prefix.length
  const active = activeCommentId.value

  if (frag.from < p) {
    return renderInlinesHtml(
      sliceForView(block.inlines, 0, Math.max(0, frag.to - p)),
      prefix,
      active,
    )
  }
  return renderInlinesHtml(sliceForView(block.inlines, frag.from - p, frag.to - p), '', active)
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
  const style: Record<string, string> = {}
  if (isFirst) style.marginTop = '0'
  if (frag.continuation) style.textIndent = '0'
  return style
}

/* -------------------------------------------------------------------------- */
/* 重排                                                                        */
/* -------------------------------------------------------------------------- */

function sameLayout(a: readonly PageLayout[], b: readonly PageLayout[]): boolean {
  if (a.length !== b.length) return false
  for (let i = 0; i < a.length; i += 1) {
    const pa = a[i]
    const pb = b[i]
    if (!pa || !pb) return false
    if (pa.sectionIndex !== pb.sectionIndex || pa.pageNumber !== pb.pageNumber) return false
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
        fa.continuation !== fb.continuation
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
  const nextPages = paginate(items, { contentHeight: contentBoxPx(spec).height })
  const changed = options.force === true || !sameLayout(pages.value, nextPages)

  if (changed) {
    viewDoc.value = cloneDoc(doc.value)
    viewNumbering.value = numberingOf(viewDoc.value)
    pages.value = nextPages
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
  if (!anchor) return
  const anchorEnd = options.anchorEnd
  void nextTick(() => {
    if (current !== token) return
    const rootEl = root.value
    if (!rootEl) return
    if (anchorEnd && (anchorEnd.blockId !== anchor.blockId || anchorEnd.offset !== anchor.offset)) {
      placeRange(rootEl, anchor, anchorEnd)
    } else {
      placeCaret(rootEl, anchor)
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

/** 按当前 DOM 重新给同一块的各片段打标（data-from / data-to），编辑后坐标才不会越用越偏 */
function retagFragments(blockId: string): void {
  const rootEl = root.value
  if (!rootEl) return
  const frags = Array.from(
    rootEl.querySelectorAll<HTMLElement>(`[data-block-id="${CSS.escape(blockId)}"]`),
  )
  let cursor = 0
  for (const frag of frags) {
    const prefix = prefixLengthOf(frag)
    const text = (frag.textContent ?? '').length
    frag.dataset.from = String(cursor)
    frag.dataset.to = String(cursor + prefix + text)
    cursor += prefix + text
  }
}

/**
 * 把页面上每一片的 DOM 读回模型。
 * 片段覆盖的区间用它自己的 data-from/data-to（分页时算出来的行边界），
 * 用 DOM 的新内容整段替换 —— 这样「删掉几个字」和「多打几个字」都能被如实记录。
 */
function syncPlain(pageEl: HTMLElement): void {
  for (const frag of Array.from(pageEl.querySelectorAll<HTMLElement>('[data-block-id]'))) {
    const blockId = frag.dataset.blockId ?? ''
    const block = findBlock(doc.value, blockId)
    if (!block) continue
    const { from, to } = fragmentRange(frag, blockId)
    replaceRange(block, from, to, readInlines(frag))
    retagFragments(blockId)
  }
}

/** 修订标记；id 在全篇唯一（Word 要求 w:ins / w:del 各自带唯一 id） */
let revSeq = -1
function mark(kind: 'ins' | 'del'): RevMark {
  let max = -1
  for (const block of doc.value.blocks) {
    if (block.t !== 'textBlock') continue
    for (const inline of block.inlines) {
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
    const block = findBlock(doc.value, blockId)
    if (!block) continue
    const { from, to } = fragmentRange(frag, blockId)
    const before = sliceStrict(block.inlines, from, to)
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
      replaceRange(block, from, to, [
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
  const range = currentRange(rootEl)
  if (!range) {
    emit('selection-change', null)
    return
  }
  const block = findBlock(doc.value, range.start.blockId)
  if (!block) {
    emit('selection-change', null)
    return
  }
  const p = prefixLength(range.start.blockId)
  const from = Math.max(0, range.start.offset - p)
  const to = Math.max(from, range.end.offset - p)
  emit('selection-change', {
    blockId: block.id,
    kind: block.kind,
    from,
    to,
    collapsed: to === from,
    bold: to > from ? rangeIsBold(block, from, to) : false,
    color: to > from ? rangeColor(block, from, to) : undefined,
  })
}

function onSelectionChange(): void {
  if (!props.editable) {
    lastCaret = null
    stickyRanges = []
    return
  }
  const rootEl = root.value
  if (!rootEl) return
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

function onBeforeInput(event: InputEvent): void {
  if (!props.editable) return
  const rootEl = root.value
  if (!rootEl) return
  const sel = document.getSelection()
  if (!sel || sel.rangeCount === 0) return
  const range = sel.getRangeAt(0)
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

function onCompositionStart(): void {
  composing = true
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

function onKeydown(event: KeyboardEvent): void {
  if (!props.editable || composing) return

  const mod = event.ctrlKey || event.metaKey
  if (mod && (event.key === 'b' || event.key === 'B')) {
    event.preventDefault()
    toggleBold()
    return
  }
  if (mod && (event.key === 'u' || event.key === 'U')) {
    event.preventDefault()
    toggleUnderline()
    return
  }
  if (mod && event.shiftKey && (event.key === 'e' || event.key === 'E')) {
    event.preventDefault()
    emit('toggle-track-changes')
    return
  }
  // ctrl+F 是浏览器查找、ctrl+G 是「查找下一个」，不 preventDefault 就抢不回来
  if (mod && !event.shiftKey && (event.key === 'f' || event.key === 'F')) {
    event.preventDefault()
    emit('open-search', 'find')
    return
  }
  if (mod && !event.shiftKey && (event.key === 'g' || event.key === 'G')) {
    event.preventDefault()
    emit('open-search', 'replace')
    return
  }
  if (mod && (event.key === 'z' || event.key === 'Z')) {
    event.preventDefault()
    if (event.shiftKey) redo()
    else undo()
    return
  }
  if (mod && (event.key === 'y' || event.key === 'Y')) {
    event.preventDefault()
    redo()
    return
  }
  // alt+4：数字改写成千分位两位小数。键盘布局不同时 alt+4 的 key 可能不是 '4'，
  // 所以两种判据都要看
  if (event.altKey && !mod && (event.key === '4' || event.code === 'Digit4')) {
    event.preventDefault()
    if (!formatSelectionAsAmount()) emit('toast', '选中内容不是有效数字')
    return
  }
  if (event.key === 'Tab') {
    // 编辑公文时 Tab 不该把焦点跳出去
    event.preventDefault()
    return
  }
  if (event.key === 'Enter') {
    event.preventDefault()
    insertParagraphBreak()
    return
  }
  if (event.key === 'Backspace' && !mod) {
    const sel = document.getSelection()
    if (!sel || !sel.isCollapsed) return
    const point = caretPoint()
    if (!point) return
    if (point.offset > prefixLength(point.blockId)) return
    const merged = mergeIntoPrevious(doc.value, point.blockId)
    if (!merged) return
    event.preventDefault()
    pushHistory(point)
    refreshLayout({
      anchor: {
        blockId: merged.blockId,
        offset: prefixLength(merged.blockId) + merged.offset,
      },
      force: true,
    })
  }
}

/** 回车：在落点切开当前块。标题类段落回车后接一个正文段（公文习惯：标题一行一段） */
function insertParagraphBreak(): void {
  const point = caretPoint()
  if (!point) return
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
  const point = caretPoint()
  if (!point) return
  pushHistory(point)

  // 纯文本粘贴：换行还原成新段落（公文里段落就是行）
  const lines = text.replace(/\r\n?/g, '\n').split('\n')
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

function setBlockKind(kind: BlockKind): void {
  const rootEl = root.value
  if (!rootEl) return
  const range = currentRange(rootEl)
  if (!range) return
  const ids = new Set(selectedRanges(rootEl).map((r) => r.blockId))
  if (ids.size === 0) ids.add(range.start.blockId)
  pushHistory()
  for (const id of ids) setBlockKindOp(doc.value, id, kind)
  refreshLayout({ anchor: range.start, anchorEnd: range.end, force: true })
}

function toggleBold(): void {
  const rootEl = root.value
  if (!rootEl) return
  const ranges = selectedRanges(rootEl)
  if (ranges.length === 0) return
  const allBold = ranges.every((range) => {
    const block = findBlock(doc.value, range.blockId)
    if (!block) return false
    const p = prefixLength(range.blockId)
    return rangeIsBold(block, Math.max(0, range.from - p), Math.max(0, range.to - p))
  })
  forSelection((model, blockId, from, to) =>
    applyFormat(model, blockId, from, to, { bold: !allBold }),
  )
}

function toggleUnderline(): void {
  const rootEl = root.value
  if (!rootEl) return
  const ranges = selectedRanges(rootEl)
  if (ranges.length === 0) return
  const allUnderline = ranges.every((range) => {
    const block = findBlock(doc.value, range.blockId)
    if (!block) return false
    const p = prefixLength(range.blockId)
    return rangeIsUnderline(block, Math.max(0, range.from - p), Math.max(0, range.to - p))
  })
  forSelection((model, blockId, from, to) =>
    applyFormat(model, blockId, from, to, { underline: !allUnderline }),
  )
}

function setColor(hex: string | null): void {
  forSelection((model, blockId, from, to) =>
    applyFormat(model, blockId, from, to, {
      color: hex ? hex.replace(/^#/, '').toUpperCase() : null,
    }),
  )
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
  const block = findBlock(doc.value, target.blockId)
  if (!block) return false
  const p = prefixLength(block.id)
  const from = Math.max(0, target.from - p)
  const to = Math.max(from, target.to - p)
  if (to <= from) return false

  const amount = formatAmount(textOf(sliceStrict(block.inlines, from, to)))
  if (amount === null) return false

  // 新数字继承原文的格式（加粗／颜色／下划线／修订标记），与「改几个字」同一路数
  const first = sliceStrict(block.inlines, from, to).find(
    (inline): inline is TextInline => inline.t === 'text',
  )
  const piece: Inline = first ? { ...first, text: amount } : { t: 'text', text: amount }

  pushHistory()
  replaceRange(block, from, to, [piece])
  // 改完仍把这串数字选中：连着按几次结果稳定（12,345.60 再解析还是它自己）
  refreshLayout({
    anchor: { blockId: block.id, offset: p + from },
    anchorEnd: { blockId: block.id, offset: p + from + amount.length },
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

  // 焦点被工具栏下拉拿走时实时选区可能已经收起来了，回退到最近记录的选区／落点
  const live = selectedRanges(rootEl)
  const target = live.length > 0 ? live[0] : stickyRanges[0]
  const caret = caretPoint()
  const blockId = target?.blockId ?? caret?.blockId
  if (blockId === undefined) return false
  const block = findBlock(doc.value, blockId)
  if (!block) return false

  const p = prefixLength(block.id)
  const from = target ? Math.max(0, target.from - p) : Math.max(0, (caret?.offset ?? 0) - p)
  const to = target ? Math.max(from, target.to - p) : from
  const char = SPECIAL_SPACES[kind]

  pushHistory(caret)
  replaceRange(block, from, to, [{ t: 'text', text: char }])
  refreshLayout({ anchor: { blockId: block.id, offset: p + from + char.length }, force: true })
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
  if (!findBlock(doc.value, blockId)) return -1
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
}

/** 点页间标记上的 × 时删掉那一枚分页符/分节符 */
function deleteBreak(blockId: string): void {
  if (!props.editable) return
  const point = caretPoint()
  pushHistory(point)
  if (!removeBreakOp(doc.value, blockId)) return
  refreshLayout({ anchor: point, force: true })
}

/** 工具栏按钮：在落点所在块之后插入分页符 / 分节符 */
function insertPageBreak(): void {
  insertBreak('page')
}

function insertSectionBreak(): void {
  insertBreak('section')
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
  refreshLayout({ force: true })
})

onBeforeUnmount(() => {
  document.removeEventListener('selectionchange', onSelectionChange)
  if (compSyncTimer !== null) clearTimeout(compSyncTimer)
})

watch([() => props.source, () => props.model], () => {
  doc.value = props.model ?? parseMd(props.source, { author: props.author })
  clearMeasureCache(cache)
  undoStack.length = 0
  redoStack.length = 0
  refreshLayout({ force: true })
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
  // 编辑层
  setBlockKind,
  toggleBold,
  toggleUnderline,
  setColor,
  formatSelectionAsAmount,
  insertSpecialSpace,
  addCommentOnSelection,
  addCommentAt,
  replyComment,
  removeComment,
  editComment,
  insertPageBreak,
  insertSectionBreak,
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
        <div class="wtp-page">
          <div
            class="wtp-content"
            :contenteditable="editable ? 'true' : undefined"
            :spellcheck="editable ? 'false' : undefined"
            @input="onInput"
            @keydown="onKeydown"
            @beforeinput="onBeforeInput"
            @paste="onPaste"
            @compositionstart="onCompositionStart"
            @compositionend="onCompositionEnd"
          >
            <div
              v-for="(frag, index) in page.fragments"
              :key="`${frag.blockId}-${frag.from}`"
              :class="`wtp-${frag.kind}`"
              :style="fragmentStyle(frag, index === 0)"
              :data-block-id="frag.blockId"
              :data-from="frag.from"
              :data-to="frag.to"
              :data-continuation="frag.continuation ? '1' : undefined"
              v-html="fragmentHtml(frag)"
            />
          </div>
          <!-- wtp-footer 提供排版（= 内置「页脚」样式），wtp-page-number 只负责定位 -->
          <div class="wtp-page-number wtp-footer">{{ page.pageNumber }}</div>
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
      <div v-if="pages.length === 0" class="wtp-page">
        <div class="wtp-content" />
        <div class="wtp-page-number wtp-footer">1</div>
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
