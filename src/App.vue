<script setup lang="ts">
import { computed, onMounted, ref } from 'vue'

import WordPaper from './components/WordPaper.vue'
import { mm } from './lib/spec'
import type { DeepPartial, Spec } from './lib/spec'

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
  '---',
  '',
  '## 附件说明',
  '',
  '本说明所附材料清单见附件一，页码自本节起重新编号。',
  '',
  '>> 江苏爱康光电科技有限公司管理人',
  '>> 2026年9月12日',
].join('\n')

const source = ref(SAMPLE)
const author = ref('张三')
const marginMm = ref(25)
const pageCount = ref(0)
const exporting = ref(false)
const paper = ref<InstanceType<typeof WordPaper> | null>(null)

/** 页边距四边一起调，用来演示规格表是可覆盖的（默认 25mm 不是写死的） */
const specOverride = computed<DeepPartial<Spec>>(() => {
  const value = mm(marginMm.value)
  return { page: { margin: { top: value, right: value, bottom: value, left: value } } }
})

async function onExport(): Promise<void> {
  exporting.value = true
  try {
    await paper.value?.downloadDocx('关于爱康光电资产核查情况的说明.docx')
  } finally {
    exporting.value = false
  }
}

// 开发期把组件实例挂到 window，供浏览器端验收脚本读取量测值对账
onMounted(() => {
  if (import.meta.env.DEV) {
    ;(window as unknown as Record<string, unknown>).__wtpPaper = paper.value
  }
})
</script>

<template>
  <div class="app">
    <header class="bar">
      <strong>WordToHtml · 公文 A4 预览</strong>
      <label class="field">
        页边距
        <input v-model.number="marginMm" type="number" min="10" max="40" step="1" />
        mm
      </label>
      <label class="field">
        修订作者
        <input v-model="author" type="text" size="6" />
      </label>
      <span class="spacer" />
      <span class="count">{{ pageCount }} 页</span>
      <button type="button" :disabled="exporting" @click="onExport">
        {{ exporting ? '导出中…' : '导出 docx' }}
      </button>
    </header>

    <main class="panes">
      <section class="pane">
        <div class="pane-head">类 md 源码</div>
        <textarea v-model="source" spellcheck="false" />
        <details class="legend">
          <summary>语法说明</summary>
          <ul>
            <li><code>#</code> 文本标题；<code>##</code>／<code>###</code>／<code>####</code> 对应 一、／（一）／1、三级标题（编号自动生成）</li>
            <li><code>@</code> 抬头（取消首行缩进）；<code>&gt;&gt;</code> 落款（右对齐）</li>
            <li><code>-</code> 列表段落；<code>!</code> 列表标题；<code>---</code> 单独一行表示分节（新起一页、页码从 1 重排）</li>
            <li><code>**文字**</code> 加粗；<code>{红|文字}</code> 改色（中文色名或 <code>#RRGGBB</code>）</li>
            <li><code>{+新增}</code> 插入修订；<code>{-删除}</code> 删除修订</li>
            <li><code>[[文字|批注内容]]</code> 批注</li>
            <li>一个非空行就是一段；空行只作分隔，不产出内容</li>
          </ul>
        </details>
      </section>

      <section class="pane">
        <div class="pane-head">A4 预览（与导出 docx 同源，改左边即时重排）</div>
        <div class="canvas">
          <WordPaper
            ref="paper"
            :source="source"
            :spec="specOverride"
            :author="author"
            @paginated="pageCount = $event"
          />
        </div>
      </section>
    </main>
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

.field {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  color: #5a5f66;
}

.field input {
  width: 56px;
  padding: 3px 6px;
  border: 1px solid #c8ccd2;
  border-radius: 4px;
  font: inherit;
}

.spacer {
  flex: 1;
}

.count {
  color: #5a5f66;
  font-variant-numeric: tabular-nums;
}

button {
  padding: 6px 14px;
  border: 1px solid #1f6feb;
  border-radius: 4px;
  background: #1f6feb;
  color: #fff;
  font: inherit;
  cursor: pointer;
}

button:disabled {
  opacity: 0.6;
  cursor: default;
}

.panes {
  display: flex;
  flex: 1;
  min-height: 0;
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

.canvas {
  flex: 1;
  min-height: 0;
  overflow: auto;
  padding: 20px;
  background: #e9eaec;
}
</style>
