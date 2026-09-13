<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'

import WordPaper from './components/WordPaper.vue'
import { toMd } from './lib/md/serialize'
import { BLOCK_KINDS, DOC_TEMPLATES, ptToPx, resolveSpec } from './lib/spec'
import type { BlockKind, DeepPartial, Spec } from './lib/spec'
import type { EditorSelection } from './lib/edit/model'

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

const specOverride = computed<DeepPartial<Spec>>(() => {
  const template = DOC_TEMPLATES.find((t) => t.key === templateKey.value) ?? DOC_TEMPLATES[0]
  return template.spec
})

/** 完整规格表。样式库要按每条样式自己的字体字号预览，所以这里要拿到解析后的值。 */
const spec = computed<Spec>(() => resolveSpec(specOverride.value))

const kind = computed<BlockKind>(() => selection.value?.kind ?? 'body')

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

/** 切到源码视图前，把当前模型序列化成 md —— 所见即所得改完总要看得到「它长什么样」 */
function switchMode(next: 'edit' | 'source'): void {
  if (next === mode.value) return
  if (next === 'source') {
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

    <main class="panes" :class="{ single: mode === 'edit' }">
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
            <li><code>[[文字|批注内容]]</code> 批注</li>
            <li>一个非空行就是一段；空行只作分隔，不产出内容</li>
          </ul>
        </details>
      </section>

      <section class="pane">
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
          />
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
  .pane-head,
  .legend,
  textarea,
  .toast {
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
