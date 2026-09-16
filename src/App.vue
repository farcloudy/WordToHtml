<script setup lang="ts">
/**
 * demo：组件（`WtpEditor`）+「类 md 源码」pane + 模式切换。
 *
 * 源码 pane 与模式切换**不是组件的一部分**（组件边界只到「纸张」，见 PLAN 13.3），
 * 所以它们留在这里，靠组件给的两个插槽挂回同一个外壳里：
 *   · `#bar-extra` → 顶栏那排模式按钮
 *   · 默认插槽    → 纸张左侧的源码 pane
 *
 * 这里同时演示「props 进 / emits 出」这条回路怎么接，以及 `content` 留空的用法。
 */
import { nextTick, onMounted, ref, shallowRef, watch } from 'vue'

import WtpEditor from './components/WtpEditor.vue'
import { replyComment } from './lib/edit/model'
import { parseMd } from './lib/md/parse'
import { toMd } from './lib/md/serialize'
import { DOC_TEMPLATES } from './lib/spec'
import { emptyDoc } from './lib/types'
import type { DocModel } from './lib/types'
import type { ShortcutOverrides } from './lib/edit/shortcuts'

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
  // 分节符后跟 kwargs 描述**它开启的那一节**：这一节独立设页脚、页码从 1 重排
  // （= W5 之前 `---` 的旧含义，写全了才是显式的）。不写 kwargs 就是全默认 ——
  // 新节关联前节，页码接着往下数。
  '--- link=off restart=on',
  '',
  '## 附件说明',
  '',
  '本说明所附材料清单见附件一，页码自本节起重新编号。',
  '',
  '>> 江苏爱康光电科技有限公司管理人',
  '>> 2026年9月12日',
].join('\n')

/**
 * 浏览器验收用的入口：`?shortcuts=bold:ctrl+shift+b,repeat:f9` 覆盖默认快捷键表。
 *
 * 只有 demo 用得上 —— 真实使用方直接把 `shortcuts` 这个 prop 递给组件。做成 URL 参数
 * 是因为验收脚本要在**首次渲染之前**把表交进去；未知动作名照原样透传（组件会忽略 + warn，
 * 那一幕本身也要验）。一个动作一项，`:` 分隔动作名与组合键，`+` 是组合键自己的分隔符。
 */
function shortcutsFromUrl(search: string): ShortcutOverrides | undefined {
  /*
   * 自己切 query，**不用 URLSearchParams**：后者按表单编码把 `+` 解成空格，
   * 而组合键的修饰键分隔符正是 `+`（`bold:ctrl+shift+b` 会被解成「ctrl shift b」，
   * 于是整张表静默退回默认值）。`%2B` 走 decodeURIComponent 也能还原成 `+`，两种写法都认。
   */
  for (const field of search.replace(/^\?/, '').split('&')) {
    if (!field.startsWith('shortcuts=')) continue
    let raw = ''
    try {
      raw = decodeURIComponent(field.slice('shortcuts='.length))
    } catch {
      return undefined
    }
    const out: Record<string, string> = {}
    for (const pair of raw.split(',')) {
      const at = pair.indexOf(':')
      if (at < 0) continue
      const action = pair.slice(0, at).trim()
      const combo = pair.slice(at + 1).trim()
      if (action === '' || combo === '') continue
      out[action] = combo
    }
    return Object.keys(out).length > 0 ? (out as ShortcutOverrides) : undefined
  }
  return undefined
}

/** `?empty=1`：演示「`content` 留空 = 空字符串」这条用法（也用来验「空内容不报错」） */
const emptyDemo = window.location.search.includes('empty=1')

/**
 * `?model=empty|rich|rich-md`：演示「直接给一份现成的模型」这条入口
 * （`model` 与 `content` 二选一，`model` 优先）。
 *
 *   · `empty`   —— `emptyDoc()`：零块文档（载入时应被补成一个空白正文段落）；
 *   · `rich`    —— 现造一份模型：修订带作者与时间戳、批注带作者与回复线程；
 *   · `rich-md` —— 上面那份内容的 **md 形态**改走 `content` 通路，当 `rich` 的对照：
 *                 md 语法里没有「修订的作者/时间戳」与「批注的回复线程」的位置，走它会被抹平。
 *
 * 前两个都**同时**递一份非空 `content`：模型被采纳时版面上是模型的内容，不是 SAMPLE ——
 * 「model 优先于 content」因此在版面上看得见。
 */
const modelKey = (() => {
  const hit = /(?:^|&)model=([^&]*)/.exec(window.location.search.replace(/^\?/, '&'))
  const key = hit ? decodeURIComponent(hit[1] ?? '') : ''
  return key === 'empty' || key === 'rich' || key === 'rich-md' ? key : ''
})()

/** 造一份「md 装不下」的模型（构造一律走 lib 既有导出，字段名不另编） */
function richModel(): DocModel {
  const doc = parseMd(
    [
      '# 模型通路演示',
      '',
      '这是{+新增的一句}，另有一处[[批注锚定的文字|批注内容]]。',
    ].join('\n'),
    { author: '李四', now: () => new Date('2026-01-02T03:04:05.000Z') },
  )
  // 回复线程只有模型（与 docx）有：md 里没有「回复某人」这个语法
  const first = doc.comments[0]
  if (first) replyComment(doc, first.id, '已核对，同意这条。', '王五', '2026-01-03T04:05:06.000Z')
  return doc
}

const presetDoc = modelKey === 'rich' || modelKey === 'rich-md' ? richModel() : null

/**
 * 走 `model` prop 递进组件的那份模型（`?model=` 之外一律 undefined，交给 `content` 通路）。
 * 必须 shallowRef：ref 会把整份模型深度转成响应式代理，而它是要交给编辑层直接改的
 * 「导出的真相」—— 代理会把 `id` 之类的内部约定搅乱，也会白白拖慢每次读写。
 */
const modelProp = shallowRef<DocModel | undefined>(
  modelKey === 'empty' ? emptyDoc() : modelKey === 'rich' ? (presetDoc ?? undefined) : undefined,
)

/**
 * `?template=govDoc`：把某个文件模板 key 当 **prop** 递进去（下拉那一路是受控回写，
 * 这一路验的是「初始 prop 就决定了版心几何」）。认不出的 key 交回第一套模板。
 */
const templateFromUrl = (() => {
  const hit = /(?:^|&)template=([^&]*)/.exec(window.location.search.replace(/^\?/, '&'))
  const key = hit ? decodeURIComponent(hit[1] ?? '') : ''
  return DOC_TEMPLATES.some((t) => t.key === key) ? key : DOC_TEMPLATES[0].key
})()

const shortcutsProp = shortcutsFromUrl(window.location.search)

/** edit = 直接在 A4 版面上写；source = 类 md 源码（demo 专属，不属于组件） */
const mode = ref<'edit' | 'source'>('edit')
const source = ref(
  modelKey === 'rich-md' && presetDoc ? toMd(presetDoc) : emptyDemo ? '' : SAMPLE,
)
const author = ref('张三')
const templateKey = ref(templateFromUrl)

/** 文件名（顶栏那一行，就是个可编辑的标题）。空内容时也不编一个假标题出来 */
const fileName = ref(emptyDemo ? '' : (SAMPLE.split('\n')[0] ?? '').replace(/^#\s*/, ''))

/** 组件报上来的页数，只用来决定「只读 md 镜像」什么时候刷新 */
const pageCount = ref(0)
const mdView = ref('')
const editor = ref<InstanceType<typeof WtpEditor> | null>(null)

/**
 * 演示「props → 组件 → emits」这条回路的落地：把组件回传的值攒到一个浏览器脚本读得到的
 * 地方（验收脚本靠它验受控回写与 `save_*`）。真实使用方直接把自己的状态绑上去即可，
 * 不必这么写 —— 见 PLAN.md 的「组件接口」。
 */
const probe = {
  fileName: fileName.value,
  author: author.value,
  template: templateKey.value,
  lastSaveMd: '',
  docxCount: 0,
  /**
   * 验收脚本专用：把一份现造的模型整体换进版面（验「换 model 时插入符按块 id 落回来」）。
   * 真实使用方直接改自己递给 `model` 的那份状态即可，不必这么写。
   */
  setModel(next: DocModel | undefined): void {
    modelProp.value = next
  },
}

function onUpdateFileName(value: string): void {
  fileName.value = value
  probe.fileName = value
}

function onUpdateAuthor(value: string): void {
  author.value = value
  probe.author = value
}

function onUpdateTemplate(value: string): void {
  templateKey.value = value
  probe.template = value
}

function onSaveMd(md: string): void {
  probe.lastSaveMd = md
}

function onSaveDocx(): void {
  probe.docxCount += 1
}

function onPaginated(count: number): void {
  pageCount.value = count
}

/** 切到源码视图前，把当前模型序列化成 md —— 所见即所得改完总要看得到「它长什么样」 */
function switchMode(next: 'edit' | 'source'): void {
  if (next === mode.value) return
  if (next === 'source') {
    source.value = editor.value?.toMd() ?? source.value
    mdView.value = source.value
  }
  mode.value = next
}

/** 源码视图跟着模型走（只读镜像），编辑结果随时能看到它的 md 形态 */
function refreshMdView(): void {
  mdView.value = editor.value?.toMd() ?? ''
}

watch(
  () => [pageCount.value, mode.value],
  () => {
    if (mode.value === 'source') void nextTick(refreshMdView)
  },
)

onMounted(() => {
  // 开发期钩子：验收脚本照旧按 window.__wtpPaper 读量测、调方法（组件把实例交了出来）
  if (import.meta.env.DEV) {
    ;(window as unknown as Record<string, unknown>).__wtpPaper = editor.value?.getPaper() ?? null
  }
  ;(window as unknown as Record<string, unknown>).__wtpDemo = probe
})
</script>

<template>
  <WtpEditor
    ref="editor"
    :content="source"
    :model="modelProp"
    :file-name="fileName"
    :author="author"
    :template="templateKey"
    :shortcuts="shortcutsProp"
    :editable="mode === 'edit'"
    @paginated="onPaginated"
    @save_md="onSaveMd"
    @save_docx="onSaveDocx"
    @update:file-name="onUpdateFileName"
    @update:author="onUpdateAuthor"
    @update:template="onUpdateTemplate"
  >
    <!-- 模式切换挂在组件的顶栏插槽里（它不属于组件） -->
    <template #bar-extra>
      <div class="tabs">
        <button type="button" :class="{ 'is-on': mode === 'edit' }" @click="switchMode('edit')">
          所见即所得
        </button>
        <button type="button" :class="{ 'is-on': mode === 'source' }" @click="switchMode('source')">
          类 md 源码
        </button>
      </div>
    </template>

    <!-- 类 md 源码 pane 挂在纸张左侧（同样不属于组件） -->
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
      <details class="legend md-mirror">
        <summary>当前模型的 md 形态（只读镜像）</summary>
        <pre>{{ mdView }}</pre>
      </details>
    </section>
  </WtpEditor>
</template>

<style scoped>
/*
 * demo 自己的样式：模式按钮、源码 pane、语法说明与只读镜像。
 * 纸与外壳的样式在 WtpEditor.vue 里；这里重复几条通用类（.pane / .pane-head）是因为
 * 插槽里的节点带的是**使用方**的 scope，组件那侧的 scoped 样式选不到它们。
 */
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

/*
 * 打印：demo 只负责把**自己**这些元素藏掉（源码 pane 的三个件）。
 * 外壳（顶栏、功能区、导航窗格、浮动面板）与纸张容器的样式在组件里，那边自己会藏。
 */
@media print {
  .pane-head,
  .legend,
  textarea {
    display: none !important;
  }

  .pane {
    display: block;
    overflow: visible;
  }

  .pane + .pane {
    border-left: 0;
  }
}
</style>
