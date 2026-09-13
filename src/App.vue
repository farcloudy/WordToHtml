<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import type { Ref } from 'vue'

import WordPaper from './components/WordPaper.vue'
import { toMd } from './lib/md/serialize'
import { BLOCK_KINDS, DOC_TEMPLATES, ptToPx, resolveSpec } from './lib/spec'
import type { Align, BlockKind, DeepPartial, Spec } from './lib/spec'
import type { CellVerticalAlign } from './lib/types'
import type { EditorSelection } from './lib/edit/model'
import type { OutlineEntry } from './lib/edit/outline'
import type { SearchScope } from './lib/edit/search'

const SAMPLE = [
  '# 关于爱康光电资产核查情况的说明',
  '',
  '@ 苏州市公安局经济犯罪侦查支队：',
  '',
  '我方于2026年9月1日收到你单位《调取证据通知书》（苏公经侦调字〔2026〕第37号），现就通知书所列事项说明如下。',
  '',
  '## 债务人基本情况',
  '',
  '债务人爱康光电科技有限公司成立于2015年3月，注册资本人民币5000万元，登记住所为苏州市工业园区星湖街328号。',
  '',
  '## 资产核查情况',
  '',
  '经核查，债务人名下资产可分为不动产、机器设备与对外投资三类。其中{+债务人已于2025年12月将部分设备对外处置}，**相关合同原件已随本说明一并提供**。',
  '',
  '### 不动产',
  '',
  '#### 已办理抵押登记的部分',
  '',
  '- 位于苏州市工业园区星湖街328号的厂房一处，建筑面积约4200平方米',
  '- 位于苏州市吴中区东吴北路的办公用房一处，建筑面积约860平方米',
  '',
  '! 上述两处不动产均已办理抵押登记，抵押权人为中国工商银行苏州分行',
  '',
  '{-债务人陈述其名下另有车位一处，经核查未在登记簿中查到相应记载}。',
  '',
  '### 机器设备',
  '',
  '债务人申报的生产设备共47台（套），管理人已现场清点，{红|其中3台因搬迁灭失}[[其中3台|灭失设备的名称与型号需另行说明]]，其余设备存放于厂区内。',
  '',
  ':::table minLines=2',
  '> 单位：元',
  '| 设备名称 | 数量 | 账面原值 |',
  '| 数控加工中心 | 6 | 1,860,000.00 |',
  '| 注塑机{br}（含配套模具） | 11 | 2,340,000.00 |',
  '| 检测仪器 | 9 | 415,000.00 |',
  '< 注：上列账面原值取自固定资产明细账，未经审计',
  ':::',
  '',
  '关于对外投资部分，债务人申报其持有苏州爱康新材料有限公司30%股权、苏州爱康智能装备有限公司18%股权以及江苏爱康光电研究所有限公司10%股权。经调取市场监管部门登记信息核对，上述三家公司均处于存续状态，其中苏州爱康新材料有限公司已于2025年8月被列入经营异常名录，苏州爱康智能装备有限公司的注册资本尚未实缴到位。管理人已分别向上述三家公司发出书面通知，要求其提供最近三年经审计的财务报表、公司章程、股东会决议以及股权质押情况说明；截至本说明出具之日，仅苏州爱康智能装备有限公司回函表示正在整理资料，其余两家公司未予回复。鉴于股权价值评估须以财务报表为基础，管理人拟在收齐上述材料后另行委托评估机构对股权价值进行评估，评估结论将作为后续财产变价方案的依据。此外，债务人陈述其曾于2024年6月与第三方签订股权转让框架协议，约定转让所持苏州爱康新材料有限公司全部股权，因受让方未按期支付首期款，该协议已于2025年3月解除，相关违约金债权是否已实际发生，管理人正在进一步核查。另据债务人陈述，其于2023年以设备融资租赁方式取得生产设备11台（套），租赁期限五年，截至本说明出具之日尚有租金未付，出租人已向管理人申报债权，管理人正在对该笔融资租赁的性质进行审查，即认定为融资租赁债权抑或取回权，将直接影响该批设备是否纳入债务人财产范围。关于应收账款，债务人账面记载的应收账款共37笔，金额合计人民币2860万元，其中账龄超过三年的有14笔；管理人已向全部债务人发出催收通知，截至本说明出具之日收到回函9份，回款人民币118万元，其余款项管理人将继续跟进催收。此外，管理人于2026年7月10日向全体已知债权人发出债权申报通知，申报期限至2026年8月25日止；截至本说明出具之日，共收到债权申报43笔，申报金额合计人民币1.14亿元，管理人已完成初步审查38笔，其中确认32笔、不予确认2笔、暂缓确认4笔。对于暂缓确认的债权，管理人的主要考虑是其担保物权是否已经依法设立并办理登记，以及申报人所主张的利息计算标准是否符合合同约定与法律规定，管理人将在补充核查后另行出具审查意见。同时，管理人对债务人涉诉及执行情况进行了梳理，债务人作为被告的民事诉讼案件共计6件，涉案金额约人民币3200万元，其中2件已经一审判决、1件处于二审阶段、3件尚未开庭；债务人作为被执行人的执行案件共计4件，执行标的合计约人民币900万元，管理人已向各执行法院提交了中止执行的申请。上述诉讼与执行情况可能对债务人财产范围及债权数额产生影响，管理人将持续跟进并及时向你单位报告。',
  '',
  '关于债务人财产的变价方案，管理人已依照《中华人民共和国企业破产法》第一百一十二条的规定拟定初步意见，并提交第一次债权人会议审议。会议召开前，管理人已将变价方案的征求意见稿送达全体已知债权人，其中三名债权人提出了书面异议，异议主要集中在评估基准日的选取以及部分机器设备是否应当与厂房一并处置两个方面。管理人经研究认为，评估基准日以人民法院受理破产申请之日为宜，理由是该日的资产状况有完整的财务凭证与现场盘点记录可供核对；至于机器设备与厂房是否应当一并处置，管理人倾向于分别处置，因为厂房的抵押权人已明确表示愿意单独受让，若将两者合并处置，可能导致抵押权人的优先受偿顺位与其他债权人的利益发生冲突，且不利于财产的及时变价。此外，管理人已就评估机构的选聘事项征询了债权人委员会的意见，拟从人民法院管理人名册中随机确定三家候选机构，再以书面询价的方式确定受托机构，评估费用列入破产费用优先支付；评估报告出具后，管理人将把评估结论与变价方案一并提交债权人会议表决，并根据会议决议另行公告拍卖或者变卖的安排。上述评估与变价工作预计在债权申报期届满后两个月内完成，管理人会按期向你单位书面报告进展；如发现债务人有转移财产或者其他损害债权人利益的行为，管理人将及时提请人民法院处理。',
  '',
  '---',
  '',
  '## 附件说明',
  '',
  '本说明所附材料清单见附件一，页码自本节起重新编号。',
  '',
  '>> 江苏爱康光电科技有限公司管理人',
  '>> 2026年9月12日',
].join('\n')

const KIND_LABEL: Record<BlockKind, string> = {
  title: '标题',
  h1: '一级标题（一、）',
  h2: '二级标题（（一））',
  h3: '三级标题（1、）',
  body: '正文',
  salutation: '抬头',
  signature: '落款',
  attachment: '附件',
  listTitle: '列表标题',
  listItem: '列表段落',
}

const COLORS: { label: string; value: string }[] = [
  { label: '红', value: 'FF0000' },
  { label: '黑', value: '000000' },
  { label: '蓝', value: '0000FF' },
  { label: '绿', value: '008000' },
  { label: '紫', value: '800080' },
  { label: '橙', value: 'FF8C00' },
  { label: '灰', value: '808080' },
]

/**
 * 可插入的特殊空格。三个码点的宽度是排版意义上的（全宽/半宽/四分之一），
 * 与字体无关；普通空格会被 HTML 折叠，这三个不会。
 */
const SPACES = [
  { kind: 'em', label: '全宽空格（U+2003）' },
  { kind: 'en', label: '半宽空格（U+2002）' },
  { kind: 'quarterEm', label: '四分之一宽空格（U+2005）' },
] as const

/**
 * 表格格内的两组对齐按钮。水平只做左 / 居中 / 右三档（不提供两端对齐），
 * 垂直三档，缺省是顶端 —— 按钮的 active 态吃 selection.table 里已解析默认值的 alignH / alignV。
 */
const TABLE_H_ALIGNS: { value: Align; label: string; title: string }[] = [
  { value: 'left', label: '左', title: '格内水平左对齐' },
  { value: 'center', label: '居中', title: '格内水平居中' },
  { value: 'right', label: '右', title: '格内水平右对齐' },
]
const TABLE_V_ALIGNS: { value: CellVerticalAlign; label: string; title: string }[] = [
  { value: 'top', label: '顶端', title: '格内顶端对齐' },
  { value: 'middle', label: '居中', title: '格内垂直居中' },
  { value: 'bottom', label: '底端', title: '格内底端对齐' },
]

/** edit = 直接在 A4 版面上写（面向用户）；source = 类 md 源码（给开发/排错用） */
const mode = ref<'edit' | 'source'>('edit')
const source = ref(SAMPLE)
const author = ref('张三')
/** 当前文件模板（样式 + 页边距是一体的，所以只有一个下拉）。默认第一套，与既有行为一致 */
const templateKey = ref(DOC_TEMPLATES[0].key)
const trackChanges = ref(false)
const pageCount = ref(0)
const exporting = ref(false)
const paper = ref<InstanceType<typeof WordPaper> | null>(null)
const selection = ref<EditorSelection | null>(null)
const commentDraft = ref('')
const mdView = ref('')
const toastText = ref('')
/** 提示条的计时器；关掉页面时也要清掉，别留一个 setTimeout 在那里 */
let toastTimer: ReturnType<typeof setTimeout> | null = null

/** 导航窗格：大纲由组件按渲染快照推来（有标题才会出现），开关由 App 持有 */
const outline = ref<OutlineEntry[]>([])
const navOpen = ref(true)

/* 查找替换面板。位置与开关都握在 App 手里，组件只管匹配与高亮。 */
const searchOpen = ref(false)
const searchMode = ref<'find' | 'replace'>('find')
const searchQuery = ref('')
const searchReplace = ref('')
const searchRegex = ref(false)
const searchScopeMode = ref<'all' | 'selection'>('all')
/** 打开面板那一刻捕获的选区范围：焦点进面板以后实时选区就没了，只能提前记下来 */
const searchScope = ref<SearchScope[]>([])
const searchState = ref<{ total: number; current: number; error: string | null }>({
  total: 0,
  current: 0,
  error: null,
})
const findInput = ref<HTMLInputElement | null>(null)
const replaceInput = ref<HTMLInputElement | null>(null)
const previewPane = ref<HTMLElement | null>(null)
const searchPanel = ref<HTMLElement | null>(null)
/** 拖动后的位置（相对预览窗格）。null = 还没拖过，贴右上角 */
const panelPos = ref<{ x: number; y: number } | null>(null)

/** 「插入表格」面板打开时的默认规格（与 WordPaper 侧的占位默认一致） */
const TABLE_ROWS_DEFAULT = 2
const TABLE_COLS_DEFAULT = 3

/**
 * 「插入表格」面板：行数、列数由用户选，插进去的空表按这个规格生成。
 * 打开时回到默认值，免得上次调的规格在下次插入时被当成默认。
 */
const tablePanel = ref<HTMLElement | null>(null)
const tablePanelPos = ref<{ x: number; y: number } | null>(null)
const tablePanelOpen = ref(false)
const tableRows = ref(TABLE_ROWS_DEFAULT)
const tableCols = ref(TABLE_COLS_DEFAULT)
const tableRowsInput = ref<HTMLInputElement | null>(null)

/**
 * 未拖过时用各自给的贴边位置；两个分支给同一组键，免得推断出带 undefined 的联合类型。
 * 插入表格面板默认贴左上、查找面板贴右上 —— 两个浮层同时开着也不会叠在一起。
 */
function floatingPanelStyle(
  pos: { x: number; y: number } | null,
  fallback: Record<string, string>,
): Record<string, string> {
  if (!pos) return fallback
  return { left: `${pos.x}px`, top: `${pos.y}px`, right: 'auto' }
}

const searchPanelStyle = computed(() =>
  floatingPanelStyle(panelPos.value, { right: '12px', left: 'auto', top: '12px' }),
)
const tablePanelStyle = computed(() =>
  floatingPanelStyle(tablePanelPos.value, { right: 'auto', left: '12px', top: '12px' }),
)

const specOverride = computed<DeepPartial<Spec>>(() => {
  const template = DOC_TEMPLATES.find((t) => t.key === templateKey.value) ?? DOC_TEMPLATES[0]
  return template.spec
})

/** 完整规格表。样式库要按每条样式自己的字体字号预览，所以这里要拿到解析后的值。 */
const spec = computed<Spec>(() => resolveSpec(specOverride.value))

const kind = computed<BlockKind>(() => selection.value?.kind ?? 'body')

/** 光标落在表格格子里时的上下文；其余时候为 null（上下文工具条据此显示/隐藏） */
const tableCtx = computed(() => selection.value?.table ?? null)

/** 上下文工具条上的落点提示：正文行显示行列（下标 +1），表头/附注行没有列的概念 */
const tablePosLabel = computed(() => {
  const t = tableCtx.value
  if (!t) return ''
  if (t.role === 'unit') return '表头行'
  if (t.role === 'note') return '附注行'
  return `第 ${t.row + 1} 行第 ${t.col + 1} 列`
})

/**
 * 样式库按钮照 Word 的做法「所见即所得」：用这条样式自己的字体与字重显示按钮文字。
 * 字号按样式字号缩放但夹在 12–15px，否则标题那类大字号会把整条样式栏撑高一倍。
 */
function chipStyle(k: BlockKind): Record<string, string> {
  const s = spec.value.styles[k]
  const size = Math.min(15, Math.max(12, ptToPx(s.sizePt) * 0.62))
  return {
    fontFamily: `"${s.ascii}", "${s.eastAsia}", serif`,
    fontWeight: s.bold ? '700' : '400',
    fontSize: `${size.toFixed(1)}px`,
  }
}

function onSelectionChange(value: EditorSelection | null): void {
  selection.value = value
}

function insertComment(): void {
  const text = commentDraft.value.trim()
  if (text === '') return
  const id = paper.value?.addCommentOnSelection(text) ?? -1
  if (id >= 0) commentDraft.value = ''
}

/**
 * 提示条：只有一条，约 2 秒后自己消失。连续触发时重置同一个计时器而不是再起一个
 * —— 否则前一个 setTimeout 会把后一条消息提前清掉。
 */
function showToast(message: string): void {
  toastText.value = message
  if (toastTimer !== null) clearTimeout(toastTimer)
  toastTimer = setTimeout(() => {
    toastTimer = null
    toastText.value = ''
  }, 2000)
}

/**
 * 下拉选完就复位回占位项：不复位的话「再选同一项」不会再触发 change，
 * 想连插两个同宽空格就做不到。
 */
function onInsertSpace(event: Event): void {
  const select = event.target as HTMLSelectElement
  const entry = SPACES.find((s) => s.kind === select.value)
  select.value = ''
  if (!entry) return
  if (!paper.value?.insertSpecialSpace(entry.kind)) showToast('请先把插入符放到版面上')
}

/** 修订模式由 App 持有（顶栏那个复选框也绑着它），快捷键只是换个入口 */
function toggleTrackChanges(): void {
  trackChanges.value = !trackChanges.value
}

/* -------------------------------------------------------------------------- */
/* 查找与替换面板                                                              */
/* -------------------------------------------------------------------------- */

function runSearch(): void {
  paper.value?.setSearch(searchQuery.value, {
    regex: searchRegex.value,
    // 范围取「当前选中的文本」时把捕获到的区间传下去；空数组 = 没有任何区间，命中 0 处
    ...(searchScopeMode.value === 'selection' ? { scope: searchScope.value } : {}),
  })
}

/*
 * 面板里的输入与选项一变就重跑。
 * 用 watch 而不是在输入框上写 @input：那要依赖「v-model 的监听器先于 @input 跑」这个
 * 顺序细节，写错会慢一个字符。
 */
watch([searchQuery, searchRegex, searchScopeMode], () => {
  if (searchOpen.value) runSearch()
})

/** ctrl+F / ctrl+G 都开这一个面板，只是决定焦点落在哪个输入框 */
async function onOpenSearch(next: 'find' | 'replace'): Promise<void> {
  searchMode.value = next
  // 焦点马上要进面板，实时选区到那时就没了 —— 必须在这一刻把范围记下来
  searchScope.value = paper.value?.getSelectionScope() ?? []
  searchScopeMode.value = 'all'
  searchOpen.value = true
  panelPos.value = null
  await nextTick()
  const input = next === 'replace' ? replaceInput.value : findInput.value
  input?.focus()
  input?.select()
  if (searchQuery.value !== '') runSearch()
}

function closeSearch(): void {
  searchOpen.value = false
  panelPos.value = null
  paper.value?.clearSearch()
  searchState.value = { total: 0, current: 0, error: null }
  void nextTick(() => {
    // 焦点还给版面，接着敲字不用再点一次
    document.querySelector<HTMLElement>('.wtp-content[contenteditable="true"]')?.focus()
  })
}

function nextMatch(): void {
  paper.value?.nextMatch()
}

function prevMatch(): void {
  paper.value?.prevMatch()
}

function replaceOne(): void {
  paper.value?.replaceCurrent(searchReplace.value)
}

function replaceAllMatches(): void {
  paper.value?.replaceAll(searchReplace.value)
}

/** 打开「插入表格」面板：每次打开都回到默认规格，并把焦点交给行数框 */
async function openTablePanel(): Promise<void> {
  tableRows.value = TABLE_ROWS_DEFAULT
  tableCols.value = TABLE_COLS_DEFAULT
  tablePanelOpen.value = true
  tablePanelPos.value = null
  // 焦点进面板：数字立刻能敲，Esc 也才有接收者（Esc 只挂在面板的输入框上）。
  // 焦点离开正文不会丢掉插入点 —— 组件那边 lastCaret 记着上一次的落点。
  await nextTick()
  tableRowsInput.value?.focus()
  tableRowsInput.value?.select()
}

/** 取消：关面板并把焦点还给版面，接着敲字不用再点一次 */
function cancelTablePanel(): void {
  tablePanelOpen.value = false
  tablePanelPos.value = null
  void nextTick(() => {
    document.querySelector<HTMLElement>('.wtp-content[contenteditable="true"]')?.focus()
  })
}

/** 插完就关。不还焦点 —— 插入本身已经把插入符放进新表的第一格了 */
function closeTablePanel(): void {
  tablePanelOpen.value = false
  tablePanelPos.value = null
}

/** 按面板里选的行列数插入空表；越界值由 WordPaper 那边夹回合法区间 */
function confirmInsertTable(): void {
  paper.value?.insertTable(tableRows.value, tableCols.value)
  closeTablePanel()
}

/** 正在拖的面板（同一时刻只会有一个）与指针相对它左上角的偏移，保证「抓住哪儿就从哪儿拖」 */
let dragState: {
  panel: HTMLElement
  pos: Ref<{ x: number; y: number } | null>
  offset: { x: number; y: number }
} | null = null

/** 两个浮动面板共用一套拖动：起始函数只管把「自己的元素与位置引用」交进来 */
function beginPanelDrag(
  panel: HTMLElement | null,
  pos: Ref<{ x: number; y: number } | null>,
  event: PointerEvent,
): void {
  const pane = previewPane.value
  if (!pane || !panel) return
  const paneRect = pane.getBoundingClientRect()
  const rect = panel.getBoundingClientRect()
  // 由 right 定位切成 left 定位：不先钉住，第一下 pointermove 会跳半个面板宽
  pos.value = { x: rect.left - paneRect.left, y: rect.top - paneRect.top }
  dragState = { panel, pos, offset: { x: event.clientX - rect.left, y: event.clientY - rect.top } }
  window.addEventListener('pointermove', onPanelDrag)
  window.addEventListener('pointerup', endPanelDrag)
  event.preventDefault()
}

function startSearchDrag(event: PointerEvent): void {
  beginPanelDrag(searchPanel.value, panelPos, event)
}

function startTableDrag(event: PointerEvent): void {
  beginPanelDrag(tablePanel.value, tablePanelPos, event)
}

function onPanelDrag(event: PointerEvent): void {
  const pane = previewPane.value
  const drag = dragState
  if (!pane || !drag) return
  const paneRect = pane.getBoundingClientRect()
  const maxX = Math.max(0, paneRect.width - drag.panel.offsetWidth)
  const maxY = Math.max(0, paneRect.height - drag.panel.offsetHeight)
  drag.pos.value = {
    x: Math.min(Math.max(0, event.clientX - paneRect.left - drag.offset.x), maxX),
    y: Math.min(Math.max(0, event.clientY - paneRect.top - drag.offset.y), maxY),
  }
}

function endPanelDrag(): void {
  dragState = null
  window.removeEventListener('pointermove', onPanelDrag)
  window.removeEventListener('pointerup', endPanelDrag)
}

/** 切到源码视图前，把当前模型序列化成 md —— 所见即所得改完总要看得到「它长什么样」 */
function switchMode(next: 'edit' | 'source'): void {
  if (next === mode.value) return
  if (next === 'source') {
    // 源码视图用浏览器自带的查找；面板留着会和 props.source 的重新解析打架
    if (searchOpen.value) closeSearch()
    // 插入表格面板同理：它在只读预览上没有任何意义，还会盖住版面
    if (tablePanelOpen.value) closeTablePanel()
    source.value = paper.value ? toMd(paper.value.getModel()) : source.value
    mdView.value = source.value
  }
  mode.value = next
}

async function onExport(): Promise<void> {
  exporting.value = true
  try {
    await paper.value?.downloadDocx('关于爱康光电资产核查情况的说明.docx')
  } finally {
    exporting.value = false
  }
}

/** 源码视图跟着模型走（只读镜像），编辑结果随时能看到它的 md 形态 */
function refreshMdView(): void {
  mdView.value = paper.value ? toMd(paper.value.getModel()) : ''
}

watch(
  () => [pageCount.value, mode.value],
  () => {
    if (mode.value === 'source') void nextTick(refreshMdView)
  },
)

// 开发期把组件实例挂到 window，供浏览器端验收脚本读取量测值对账
onMounted(() => {
  if (import.meta.env.DEV) {
    ;(window as unknown as Record<string, unknown>).__wtpPaper = paper.value
  }
})

onBeforeUnmount(() => {
  if (toastTimer !== null) clearTimeout(toastTimer)
  endPanelDrag()
})
</script>

<template>
  <div class="app">
    <header class="bar">
      <strong>WordToHtml · 公文 A4 编辑器</strong>
      <div class="tabs">
        <button type="button" :class="{ 'is-on': mode === 'edit' }" @click="switchMode('edit')">
          所见即所得
        </button>
        <button type="button" :class="{ 'is-on': mode === 'source' }" @click="switchMode('source')">
          类 md 源码
        </button>
      </div>
      <button
        v-if="outline.length > 0"
        type="button"
        class="tool"
        :class="{ 'is-on': navOpen }"
        title="显示/隐藏左侧导航窗格"
        @click="navOpen = !navOpen"
      >
        导航
      </button>
      <label class="field">
        文件模板
        <select v-model="templateKey">
          <option v-for="t in DOC_TEMPLATES" :key="t.key" :value="t.key">
            {{ t.label }}
          </option>
        </select>
      </label>
      <label class="field">
        修订作者
        <input v-model="author" type="text" size="6" />
      </label>
      <label class="field checkbox">
        <input v-model="trackChanges" type="checkbox" />
        修订模式
      </label>
      <span class="spacer" />
      <span class="count">{{ pageCount }} 页</span>
      <button type="button" class="primary" :disabled="exporting" @click="onExport">
        {{ exporting ? '导出中…' : '导出 docx' }}
      </button>
    </header>

    <!--
      样式库。照 Word 的样子水平排成一排大按钮，按钮文字用各条样式自己的字体字号
      渲染 —— 一眼能对上是哪条样式，也不用先在下拉框里找。
    -->
    <div v-if="mode === 'edit'" class="styles">
      <span class="styles-label">样式</span>
      <button
        v-for="k in BLOCK_KINDS"
        :key="k"
        type="button"
        class="style-chip"
        :class="{ 'is-on': k === kind }"
        :style="chipStyle(k)"
        :title="KIND_LABEL[k]"
        @mousedown.prevent
        @click="paper?.setBlockKind(k)"
      >
        {{ KIND_LABEL[k] }}
      </button>
    </div>

    <div v-if="mode === 'edit'" class="toolbar">
      <button
        type="button"
        class="tool"
        :class="{ 'is-on': selection?.bold }"
        title="加粗（Ctrl+B）"
        @mousedown.prevent
        @click="paper?.toggleBold()"
      >
        <b>B</b>
      </button>

      <span class="field">
        颜色
        <span class="swatches">
          <button
            v-for="c in COLORS"
            :key="c.value"
            type="button"
            class="swatch"
            :style="{ background: `#${c.value}` }"
            :title="c.label"
            @mousedown.prevent
            @click="paper?.setColor(c.value)"
          />
          <button
            type="button"
            class="swatch clear"
            title="默认颜色"
            @mousedown.prevent
            @click="paper?.setColor(null)"
          >
            ×
          </button>
        </span>
      </span>

      <span class="sep" />

      <button
        type="button"
        class="tool"
        title="撤销（Ctrl+Z）"
        @mousedown.prevent
        @click="paper?.undo()"
      >
        ↺
      </button>
      <button
        type="button"
        class="tool"
        title="重做（Ctrl+Y）"
        @mousedown.prevent
        @click="paper?.redo()"
      >
        ↻
      </button>

      <span class="sep" />

      <button
        type="button"
        class="tool"
        title="在光标所在段落后插入分页符（只换页，页码连续）"
        @mousedown.prevent
        @click="paper?.insertPageBreak()"
      >
        分页符
      </button>
      <button
        type="button"
        class="tool"
        title="在光标所在段落后插入分节符（新起一页，页码从 1 重排）"
        @mousedown.prevent
        @click="paper?.insertSectionBreak()"
      >
        分节符
      </button>
      <button
        type="button"
        class="tool"
        title="在光标所在段落后插入一张空表格（行数、列数可选）"
        @mousedown.prevent
        @click="openTablePanel"
      >
        插入表格
      </button>

      <span class="sep" />

      <span class="field">
        插入
        <select title="在插入符处插入一个特殊空格" @change="onInsertSpace">
          <option value="">空格…</option>
          <option v-for="s in SPACES" :key="s.kind" :value="s.kind">{{ s.label }}</option>
        </select>
      </span>

      <span class="sep" />

      <span class="field comment-field">
        批注
        <input
          v-model="commentDraft"
          type="text"
          placeholder="选中文字后填写"
          @mousedown="paper?.keepSelection()"
          @keydown.enter.prevent="insertComment"
        />
        <button type="button" @mousedown.prevent @click="insertComment">添加</button>
      </span>
    </div>

    <!--
      表格的上下文工具条：只在光标落在格子里时出现（跟着 selection-change 的 table 上下文走）。
      所有按钮与 radio 都 @mousedown.prevent —— 焦点不离开正文，落点与 native 选区才保得住。
    -->
    <div v-if="mode === 'edit' && tableCtx" class="toolbar sub-toolbar">
      <span class="tk-hint">表格 · {{ tablePosLabel }}</span>

      <span class="tk-group">
        <span class="tk-label">行</span>
        <button
          type="button"
          class="tool"
          title="在光标所在行的上方插入一行"
          @mousedown.prevent
          @click="paper?.insertTableRow('above')"
        >
          上方插入行
        </button>
        <button
          type="button"
          class="tool"
          title="在光标所在行的下方插入一行"
          @mousedown.prevent
          @click="paper?.insertTableRow('below')"
        >
          下方插入行
        </button>
        <button
          type="button"
          class="tool"
          :disabled="tableCtx.role !== 'body' || tableCtx.bodyRows <= 1"
          title="删除光标所在行（只剩一个正文行、或光标在表头行／附注行时不可用）"
          @mousedown.prevent
          @click="paper?.removeTableRow()"
        >
          删除行
        </button>
      </span>

      <span class="tk-group">
        <span class="tk-label">列</span>
        <button
          type="button"
          class="tool"
          title="在光标所在列的左侧插入一列"
          @mousedown.prevent
          @click="paper?.insertTableColumn('left')"
        >
          左侧插入列
        </button>
        <button
          type="button"
          class="tool"
          title="在光标所在列的右侧插入一列"
          @mousedown.prevent
          @click="paper?.insertTableColumn('right')"
        >
          右侧插入列
        </button>
        <button
          type="button"
          class="tool"
          :disabled="tableCtx.columns <= 1"
          title="删除光标所在列（只剩一列时不可用）"
          @mousedown.prevent
          @click="paper?.removeTableColumn()"
        >
          删除列
        </button>
      </span>

      <span class="tk-group">
        <span class="tk-label">行高</span>
        <label class="tk-radio" @mousedown.prevent>
          <input
            @mousedown.prevent
            type="radio"
            name="tk-minlines"
            :checked="tableCtx.minLines === 1"
            @click="paper?.setTableMinLines(1)"
          />
          最小一行
        </label>
        <label class="tk-radio" @mousedown.prevent>
          <input
            @mousedown.prevent
            type="radio"
            name="tk-minlines"
            :checked="tableCtx.minLines === 2"
            @click="paper?.setTableMinLines(2)"
          />
          最小两行
        </label>
      </span>

      <span class="tk-group">
        <span class="tk-label">表头行</span>
        <label class="tk-radio" @mousedown.prevent>
          <input
            @mousedown.prevent
            type="radio"
            name="tk-unit"
            :checked="tableCtx.hasUnit"
            @click="paper?.setTableRoleRow('unit', true)"
          />
          有
        </label>
        <label class="tk-radio" @mousedown.prevent>
          <input
            @mousedown.prevent
            type="radio"
            name="tk-unit"
            :checked="!tableCtx.hasUnit"
            @click="paper?.setTableRoleRow('unit', false)"
          />
          无
        </label>
      </span>

      <span class="tk-group">
        <span class="tk-label">附注行</span>
        <label class="tk-radio" @mousedown.prevent>
          <input
            @mousedown.prevent
            type="radio"
            name="tk-note"
            :checked="tableCtx.hasNote"
            @click="paper?.setTableRoleRow('note', true)"
          />
          有
        </label>
        <label class="tk-radio" @mousedown.prevent>
          <input
            @mousedown.prevent
            type="radio"
            name="tk-note"
            :checked="!tableCtx.hasNote"
            @click="paper?.setTableRoleRow('note', false)"
          />
          无
        </label>
      </span>

      <span class="tk-group">
        <span class="tk-label">水平</span>
        <button
          v-for="a in TABLE_H_ALIGNS"
          :key="a.value"
          type="button"
          class="tool"
          :class="{ 'is-on': tableCtx.alignH === a.value }"
          :title="a.title"
          @mousedown.prevent
          @click="paper?.setTableCellAlignH(a.value)"
        >
          {{ a.label }}
        </button>
      </span>

      <span class="tk-group">
        <span class="tk-label">垂直</span>
        <button
          v-for="a in TABLE_V_ALIGNS"
          :key="a.value"
          type="button"
          class="tool"
          :class="{ 'is-on': tableCtx.alignV === a.value }"
          :title="a.title"
          @mousedown.prevent
          @click="paper?.setTableCellAlignV(a.value)"
        >
          {{ a.label }}
        </button>
      </span>

      <span class="tk-group tk-right">
        <button
          type="button"
          class="tool tk-danger"
          title="删除整张表格（不二次确认，可 Ctrl+Z 撤销）"
          @mousedown.prevent
          @click="paper?.removeTable()"
        >
          删除表格
        </button>
      </span>
    </div>

    <main class="panes" :class="{ single: mode === 'edit' }">
      <!-- 导航窗格：有标题才出现，挂在编辑器最左侧（不是塞在 WordPaper 里面） -->
      <section v-if="outline.length > 0 && navOpen" class="pane nav-pane">
        <div class="pane-head nav-head">
          <span>导航</span>
          <button type="button" class="nav-collapse" title="折叠导航窗格" @click="navOpen = false">
            «
          </button>
        </div>
        <ul class="nav-list">
          <li v-for="entry in outline" :key="entry.blockId" :class="`nav-lv${entry.level}`">
            <button
              type="button"
              :title="entry.prefix + entry.text"
              @click="paper?.focusBlock(entry.blockId)"
            >
              <span class="nav-prefix">{{ entry.prefix }}</span>{{ entry.text }}
            </button>
          </li>
        </ul>
      </section>

      <section v-if="mode === 'source'" class="pane">
        <div class="pane-head">类 md 源码</div>
        <textarea v-model="source" spellcheck="false" />
        <details class="legend">
          <summary>语法说明</summary>
          <ul>
            <li>
              <code>#</code> 文本标题；<code>##</code>／<code>###</code>／<code>####</code>
              对应 一、／（一）／1、三级标题（编号自动生成）
            </li>
            <li>
              <code>@</code> 抬头（取消首行缩进）；<code>&gt;&gt;</code>
              落款（右对齐）；<code>%</code> 附件标记（黑体、顶格、段后 1 行）
            </li>
            <li>
              <code>-</code> 列表段落；<code>!</code> 列表标题；<code>---</code>
              单独一行表示分节（新起一页、页码从 1 重排）
            </li>
            <li>
              <code>**文字**</code> 加粗；<code>__文字__</code> 下划线；<code>{红|文字}</code>
              改色（中文色名或 <code>#RRGGBB</code>）
            </li>
            <li><code>{+新增}</code> 插入修订；<code>{-删除}</code> 删除修订</li>
            <li><code>{br}</code> 软换行（单元格里用；零宽、不占字符位）</li>
            <li><code>:::table … :::</code> 表格围栏块（<code>&gt;</code> 单位行 / <code>&lt;</code> 附注行 / <code>|</code> 数据行）</li>
            <li><code>[[文字|批注内容]]</code> 批注</li>
            <li>一个非空行就是一段；空行只作分隔，不产出内容</li>
          </ul>
        </details>
      </section>

      <section ref="previewPane" class="pane preview-pane">
        <div v-if="mode === 'source'" class="pane-head">A4 预览（只读）</div>
        <div class="canvas">
          <WordPaper
            ref="paper"
            :source="source"
            :spec="specOverride"
            :author="author"
            :editable="mode === 'edit'"
            :track-changes="trackChanges"
            @paginated="pageCount = $event"
            @selection-change="onSelectionChange"
            @toggle-track-changes="toggleTrackChanges"
            @toast="showToast"
            @open-search="onOpenSearch"
            @search-state="searchState = $event"
            @outline-change="outline = $event"
          />
        </div>

        <!--
          查找替换面板：浮在预览区之上（该 pane 是 position:relative），拖标题栏可移动、
          拖不出窗格范围。面板本身两种功能都在，ctrl+F / ctrl+G 只决定焦点落在哪个框。
        -->
        <div v-if="searchOpen" ref="searchPanel" class="search-panel" :style="searchPanelStyle">
          <div class="search-head" @pointerdown="startSearchDrag">
            <span class="search-title">{{ searchMode === 'replace' ? '查找与替换' : '查找' }}</span>
            <button
              type="button"
              class="search-close"
              title="关闭（Esc）"
              @pointerdown.stop
              @click="closeSearch"
            >
              ×
            </button>
          </div>
          <div class="search-body">
            <label class="search-row">
              <span>查找</span>
              <input
                ref="findInput"
                v-model="searchQuery"
                type="text"
                spellcheck="false"
                @keydown.enter.exact.prevent="nextMatch"
                @keydown.enter.shift.prevent="prevMatch"
                @keydown.esc.prevent="closeSearch"
              />
            </label>
            <label class="search-row">
              <span>替换</span>
              <input
                ref="replaceInput"
                v-model="searchReplace"
                type="text"
                spellcheck="false"
                @keydown.enter.prevent="replaceOne"
                @keydown.esc.prevent="closeSearch"
              />
            </label>
            <div class="search-meta">
              <span class="search-count">{{
                searchState.total === 0
                  ? '共 0 处'
                  : `第 ${searchState.current} / 共 ${searchState.total} 处`
              }}</span>
              <span v-if="searchState.error" class="search-error">{{ searchState.error }}</span>
            </div>
            <div class="search-actions">
              <button type="button" @mousedown.prevent @click="prevMatch">上一个</button>
              <button type="button" @mousedown.prevent @click="nextMatch">下一个</button>
              <button type="button" @mousedown.prevent @click="replaceOne">替换</button>
              <button type="button" @mousedown.prevent @click="replaceAllMatches">全部替换</button>
            </div>
            <label class="search-toggle">
              <input v-model="searchRegex" name="search-regex" type="checkbox" />
              使用正则表达式
            </label>
            <div class="search-scope">
              <label>
                <input v-model="searchScopeMode" name="search-scope" type="radio" value="all" />
                全文
              </label>
              <label>
                <input
                  v-model="searchScopeMode"
                  name="search-scope"
                  type="radio"
                  value="selection"
                />
                当前选中的文本
              </label>
            </div>
          </div>
        </div>

        <!--
          「插入表格」规格面板：与查找替换面板同一套浮动卡片（可拖、×/Esc 关）。
          工具栏那个按钮只负责开它 —— 行数列数在这里定，空表按这个规格生成。
        -->
        <div v-if="tablePanelOpen" ref="tablePanel" class="table-panel" :style="tablePanelStyle">
          <div class="table-panel-head" @pointerdown="startTableDrag">
            <span class="table-panel-title">插入表格</span>
            <button
              type="button"
              class="table-panel-close"
              title="关闭（Esc）"
              @pointerdown.stop
              @click="cancelTablePanel"
            >
              ×
            </button>
          </div>
          <div class="table-panel-body">
            <label class="table-panel-row">
              <span>行数</span>
              <input
                ref="tableRowsInput"
                v-model.number="tableRows"
                type="number"
                min="1"
                max="30"
                step="1"
                @keydown.enter.prevent="confirmInsertTable"
                @keydown.esc.prevent="cancelTablePanel"
              />
            </label>
            <label class="table-panel-row">
              <span>列数</span>
              <input
                v-model.number="tableCols"
                type="number"
                min="1"
                max="12"
                step="1"
                @keydown.enter.prevent="confirmInsertTable"
                @keydown.esc.prevent="cancelTablePanel"
              />
            </label>
            <div class="table-panel-actions">
              <button type="button" @mousedown.prevent @click="confirmInsertTable">插入</button>
              <button type="button" @mousedown.prevent @click="cancelTablePanel">取消</button>
            </div>
          </div>
        </div>
        <details v-if="mode === 'source'" class="legend md-mirror">
          <summary>当前模型的 md 形态（只读镜像）</summary>
          <pre>{{ mdView }}</pre>
        </details>
      </section>
    </main>

    <!-- 提示条：无效输入、插入失败这类一句话反馈。固定定位，不占版面 -->
    <div v-if="toastText" class="toast" role="status">{{ toastText }}</div>
  </div>
</template>

<style scoped>
.app {
  display: flex;
  flex-direction: column;
  height: 100vh;
  overflow: hidden;
}

.bar {
  display: flex;
  align-items: center;
  gap: 14px;
  padding: 10px 16px;
  border-bottom: 1px solid #d8dade;
  background: #fafafa;
  font-size: 13px;
}

.bar strong {
  font-size: 14px;
}

.tabs {
  display: inline-flex;
  border: 1px solid #c8ccd2;
  border-radius: 4px;
  overflow: hidden;
}

.tabs button {
  padding: 4px 12px;
  border: 0;
  background: #fff;
  color: #4a4f56;
  font: inherit;
  cursor: pointer;
}

.tabs button.is-on {
  background: #1f6feb;
  color: #fff;
}

.field {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  color: #5a5f66;
}

.field input,
.field select {
  padding: 3px 6px;
  border: 1px solid #c8ccd2;
  border-radius: 4px;
  font: inherit;
}

.field input {
  width: 120px;
}

.field select {
  max-width: 230px;
}

.field.checkbox input {
  width: auto;
}

.spacer {
  flex: 1;
}

.count {
  color: #5a5f66;
  font-variant-numeric: tabular-nums;
}

button.primary {
  padding: 6px 14px;
  border: 1px solid #1f6feb;
  border-radius: 4px;
  background: #1f6feb;
  color: #fff;
  font: inherit;
  cursor: pointer;
}

button.primary:disabled {
  opacity: 0.6;
  cursor: default;
}

.styles {
  display: flex;
  align-items: center;
  gap: 6px;
  flex-wrap: wrap;
  padding: 8px 16px;
  border-bottom: 1px solid #e4e6ea;
  background: #f6f7f9;
}

.styles-label {
  margin-right: 2px;
  color: #8a9099;
  font-size: 12px;
}

/* 按钮文字由 inline style 按各条样式自己的字体字号渲染（见 chipStyle） */
.style-chip {
  padding: 5px 12px;
  border: 1px solid #d5d9df;
  border-radius: 5px;
  background: #fff;
  color: #33383f;
  line-height: 1.35;
  white-space: nowrap;
  cursor: pointer;
}

.style-chip:hover {
  border-color: #a9b0b8;
}

.style-chip.is-on {
  border-color: #1f6feb;
  background: #e8f0fe;
  box-shadow: inset 0 0 0 1px #1f6feb;
}

.toolbar {
  display: flex;
  align-items: center;
  gap: 10px;
  flex-wrap: wrap;
  padding: 8px 16px;
  border-bottom: 1px solid #d8dade;
  background: #fff;
  font-size: 13px;
}

/* 表格的上下文工具条：复用主工具栏的观感，次要色 + 更紧的行距，别另起一套设计 */
.sub-toolbar {
  gap: 12px;
  padding: 6px 16px;
  background: #f6f7f9;
  font-size: 12px;
}

.sub-toolbar .tool {
  padding: 2px 8px;
  font-size: 12px;
}

.sub-toolbar .tool:disabled {
  opacity: 0.45;
  cursor: not-allowed;
}

.tk-hint {
  color: #33383f;
  font-weight: 600;
}

.tk-group {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  color: #5a5f66;
  /* 换行只发生在组与组之间：组自身不收缩、不拆开（窄窗口下也不许错位） */
  flex: none;
}

/* 「删除表格」贴到最右，与行/列按钮拉开距离 */
.tk-right {
  margin-left: auto;
}

.tk-danger {
  border-color: #d8b4b0;
  color: #b3261e;
}

.tk-label {
  color: #8a9099;
}

.tk-radio {
  display: inline-flex;
  align-items: center;
  gap: 2px;
  cursor: pointer;
}

.tk-radio input {
  margin: 0;
}

.sep {
  width: 1px;
  height: 20px;
  background: #e0e2e6;
}

.tool {
  min-width: 30px;
  padding: 4px 8px;
  border: 1px solid #c8ccd2;
  border-radius: 4px;
  background: #fff;
  color: #33383f;
  font: inherit;
  cursor: pointer;
}

.tool.is-on {
  border-color: #1f6feb;
  background: #e8f0fe;
  color: #1f6feb;
}

.swatches {
  display: inline-flex;
  gap: 4px;
}

.swatch {
  width: 20px;
  height: 20px;
  padding: 0;
  border: 1px solid #c8ccd2;
  border-radius: 3px;
  cursor: pointer;
}

.swatch.clear {
  background: #fff;
  color: #8a9099;
  line-height: 1;
}

.comment-field input {
  width: 200px;
}

.comment-field button {
  padding: 4px 10px;
  border: 1px solid #c8ccd2;
  border-radius: 4px;
  background: #fff;
  font: inherit;
  cursor: pointer;
}

.panes {
  display: flex;
  flex: 1;
  min-height: 0;
}

.panes.single .pane {
  justify-content: flex-start;
}

.pane {
  display: flex;
  flex-direction: column;
  min-width: 0;
  flex: 1;
}

.pane + .pane {
  border-left: 1px solid #d8dade;
}

/* 预览窗格是浮动查找面板的定位父级 */
.preview-pane {
  position: relative;
}

/*
 * 导航窗格固定 200px。这个宽度不是随便定的：预览区还要放得下两页并排的 A4
 * （约 1605px），窗格再宽就会把并排挤掉（verify:p2 直接验这件事）。
 */
.nav-pane {
  flex: 0 0 200px;
  background: #fbfbfc;
}
.nav-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 6px;
}
.nav-collapse {
  padding: 0 6px;
  border: 1px solid #c8ccd2;
  border-radius: 4px;
  background: #fff;
  color: #5a5f66;
  font: inherit;
  line-height: 1.4;
  cursor: pointer;
}
.nav-collapse:hover {
  border-color: #8a9099;
}
.nav-list {
  flex: 1;
  margin: 0;
  padding: 6px 0;
  overflow: auto;
  list-style: none;
}
.nav-list li button {
  display: block;
  width: 100%;
  padding: 4px 10px;
  border: 0;
  background: transparent;
  color: #33383f;
  font: inherit;
  font-size: 12px;
  line-height: 1.5;
  text-align: left;
  cursor: pointer;
}
.nav-list li button:hover {
  background: #eef2f8;
}
.nav-lv2 button {
  padding-left: 22px;
}
.nav-lv3 button {
  padding-left: 34px;
}
.nav-prefix {
  color: #8a9099;
}

/* 浮动面板的公共卡片外观与内件：查找替换面板与插入表格面板共用同一套。
   拆成两套类名而不用一个基类，是为了让「谁是哪个面板」在模板里一眼可辨；
   样式靠分组选择器共享，不重复声明。 */
.search-panel,
.table-panel {
  position: absolute;
  z-index: 12;
  width: 272px;
  border: 1px solid #c8ccd2;
  border-radius: 8px;
  background: #fff;
  box-shadow: 0 6px 20px rgba(20, 24, 30, 0.18);
  font-size: 12px;
  color: #33383f;
}
/* 插入表格面板只有两个数字输入，比查找面板窄 */
.table-panel {
  width: 186px;
}
.search-head,
.table-panel-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 6px 10px;
  border-bottom: 1px solid #e4e6ea;
  border-radius: 7px 7px 0 0;
  background: #f6f7f9;
  cursor: move;
  user-select: none;
}
.search-title,
.table-panel-title {
  font-weight: 600;
}
.search-close,
.table-panel-close {
  padding: 0 4px;
  border: 0;
  background: transparent;
  color: #8a9099;
  font-size: 16px;
  line-height: 1;
  cursor: pointer;
}
.search-body,
.table-panel-body {
  display: flex;
  flex-direction: column;
  gap: 6px;
  padding: 8px 10px 10px;
}
.search-row,
.table-panel-row {
  display: flex;
  align-items: center;
  gap: 6px;
}
.search-row span,
.table-panel-row span {
  flex: 0 0 28px;
  color: #5a5f66;
}
.search-row input,
.table-panel-row input {
  flex: 1;
  min-width: 0;
  padding: 3px 6px;
  border: 1px solid #c8ccd2;
  border-radius: 4px;
  font: inherit;
}
.search-meta {
  display: flex;
  align-items: center;
  gap: 8px;
  min-height: 18px;
}
.search-count {
  color: #5a5f66;
  font-variant-numeric: tabular-nums;
}
.search-error {
  color: #b3261e;
}
.search-actions,
.table-panel-actions {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
}
.search-actions button,
.table-panel-actions button {
  padding: 3px 8px;
  border: 1px solid #c8ccd2;
  border-radius: 4px;
  background: #fff;
  font: inherit;
  cursor: pointer;
}
.search-actions button:hover,
.table-panel-actions button:hover {
  border-color: #8a9099;
}
.search-toggle,
.search-scope label {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  color: #5a5f66;
}
.search-scope {
  display: flex;
  gap: 12px;
}

.pane-head {
  padding: 6px 12px;
  border-bottom: 1px solid #e4e6ea;
  background: #f6f7f9;
  color: #5a5f66;
  font-size: 12px;
}

textarea {
  flex: 1;
  padding: 12px 14px;
  border: 0;
  outline: none;
  resize: none;
  font-family: Consolas, 'Courier New', monospace;
  font-size: 13px;
  line-height: 1.7;
  tab-size: 2;
}

.legend {
  border-top: 1px solid #e4e6ea;
  background: #fbfbfc;
  font-size: 12px;
  color: #4a4f56;
}

.legend summary {
  padding: 7px 12px;
  cursor: pointer;
}

.legend ul {
  margin: 0;
  padding: 0 12px 10px 30px;
  line-height: 1.9;
}

.legend code {
  padding: 1px 4px;
  border-radius: 3px;
  background: #eef0f3;
  font-family: Consolas, 'Courier New', monospace;
}

.md-mirror pre {
  max-height: 220px;
  margin: 0;
  padding: 0 14px 12px;
  overflow: auto;
  font-family: Consolas, 'Courier New', monospace;
  font-size: 12px;
  line-height: 1.7;
  white-space: pre-wrap;
}

.canvas {
  flex: 1;
  min-height: 0;
  overflow: auto;
  padding: 20px;
  background: #e9eaec;
}

.toast {
  position: fixed;
  left: 50%;
  bottom: 28px;
  z-index: 20;
  padding: 8px 16px;
  border-radius: 6px;
  background: rgba(35, 39, 45, 0.92);
  color: #fff;
  font-size: 13px;
  /* 纯提示，别挡住下面的点击 */
  pointer-events: none;
  transform: translateX(-50%);
}

/*
 * 打印：只出 WordPaper 渲染的那几张 A4 纸。
 *
 * 纸张尺寸与「页面上不留白边」由 lib/render/css.ts 注入的 @page 负责（那边的尺寸来自
 * 规格表，不能在这里写死）；这里只隐藏编辑器外壳。App 的样式是 scoped 的，
 * 能选中自己的顶栏/工具栏/源码 pane，但选不中子组件内部的 .wtp-* 节点 ——
 * 那部分（批注侧栏、页间换页标记）交给组件自己的打印样式。
 */
@media print {
  .bar,
  .styles,
  .toolbar,
  .sub-toolbar,
  .pane-head,
  .legend,
  textarea,
  .toast,
  .nav-pane,
  .search-panel,
  .table-panel {
    display: none !important;
  }

  /* 屏幕上这些容器都靠固定高度 + overflow 撑出滚动区，打印时必须放开，
     否则只会印出第一屏、后面几页被裁掉 */
  .app {
    display: block;
    height: auto;
    overflow: visible;
  }

  .panes,
  .pane,
  .canvas {
    display: block;
    overflow: visible;
  }

  .pane + .pane {
    border-left: 0;
  }

  .canvas {
    padding: 0;
    background: #fff;
  }
}
</style>
