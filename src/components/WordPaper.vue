<script setup lang="ts">
/**
 * A4 分页预览。
 *
 * 这一阶段（P2）是只读的：源码进，A4 分页版面出，可导出 docx。
 * 分页不是 CSS 断页，而是「实测行盒 → 装箱 → 按行切开」算出来的，
 * 这样页码、分节重编号、跨页断行都能和 Word 对齐。
 */

import { computed, nextTick, onMounted, ref, shallowRef, watch } from 'vue'

import { toBlob } from '../lib/docx/export'
import { computeNumbering } from '../lib/numbering'
import { parseMd } from '../lib/md/parse'
import { injectCss } from '../lib/render/css'
import { renderInlinesHtml } from '../lib/render/html'
import { measureDocument } from '../lib/render/measure'
import { paginate } from '../lib/render/paginate'
import type { MeasuredItem, PageFragment, PageLayout } from '../lib/render/paginate'
import { contentBoxPx, resolveSpec } from '../lib/spec'
import type { DeepPartial, Spec } from '../lib/spec'
import { sliceInlines } from '../lib/types'
import type { DocModel, TextBlock } from '../lib/types'

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
  }>(),
  { source: '', author: '管理员' },
)

const emit = defineEmits<{ paginated: [count: number] }>()

const resolved = computed<Spec>(() => resolveSpec(props.spec))
const parsed = computed<DocModel>(() =>
  props.model ?? parseMd(props.source, { author: props.author }),
)

const host = ref<HTMLElement | null>(null)
const pages = shallowRef<PageLayout[]>([])
/** 最近一次分页实际用到的量测值，排错时用来和渲染结果对账 */
const measured = shallowRef<MeasuredItem[]>([])

/** 分页切出来的片段要能找回源块 */
const blocksById = computed(() => {
  const map = new Map<string, TextBlock>()
  for (const block of parsed.value.blocks) {
    if (block.t === 'textBlock') map.set(block.id, block)
  }
  return map
})

const numbering = computed(() =>
  computeNumbering(parsed.value.blocks, (block) =>
    block.t === 'textBlock' ? resolved.value.styles[block.kind].numbering : 'none',
  ),
)

/**
 * 渲染一个分页片段。
 * 片段偏移用的是「显示文字」坐标系（含自动编号前缀），模型里不含前缀，
 * 所以要先把前缀长度扣掉；只有首片才带前缀。
 */
function fragmentHtml(frag: PageFragment): string {
  const block = blocksById.value.get(frag.blockId)
  if (!block) return ''
  const prefix = numbering.value.get(frag.blockId) ?? ''
  const p = prefix.length

  if (frag.from < p) {
    return renderInlinesHtml(sliceInlines(block.inlines, 0, Math.max(0, frag.to - p)), prefix)
  }
  return renderInlinesHtml(sliceInlines(block.inlines, frag.from - p, frag.to - p))
}

/** 页首那一块不能带段前距，续排块还要去掉首行缩进 —— 这两点分页时已经按此记账 */
function fragmentStyle(frag: PageFragment, isFirst: boolean): Record<string, string> {
  const style: Record<string, string> = {}
  if (isFirst) style.marginTop = '0'
  if (frag.continuation) style.textIndent = '0'
  return style
}

let token = 0

async function repaginate(): Promise<void> {
  const current = ++token
  injectCss(resolved.value)
  await nextTick()
  const el = host.value
  if (!el || current !== token) return

  const spec = resolved.value
  const items = measureDocument(parsed.value, spec, el)
  if (current !== token) return
  measured.value = items
  pages.value = paginate(items, { contentHeight: contentBoxPx(spec).height })
  emit('paginated', pages.value.length)
}

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
  await repaginate()
})

watch(
  [() => props.source, () => props.model, () => props.spec],
  () => {
    void repaginate()
  },
  { deep: true },
)

async function exportDocx(): Promise<Blob> {
  return toBlob(parsed.value, resolved.value, { title: '公文' })
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
  /** 当前解析出的模型，供调用方做进一步处理 */
  getModel: (): DocModel => parsed.value,
  /** 当前生效的完整规格表 */
  getSpec: (): Spec => resolved.value,
  /** 最近一次分页用到的量测值（行数、行高、段距），排错用 */
  getMeasurements: (): MeasuredItem[] => measured.value,
  pageCount: (): number => pages.value.length,
})
</script>

<template>
  <div class="wtp-root">
    <div class="wtp-pages">
      <div
        v-for="page in pages"
        :key="`${page.sectionIndex}-${page.pageNumber}`"
        class="wtp-page"
      >
        <div class="wtp-content">
          <div
            v-for="(frag, index) in page.fragments"
            :key="`${frag.blockId}-${frag.from}`"
            :class="`wtp-${frag.kind}`"
            :style="fragmentStyle(frag, index === 0)"
            :data-block-id="frag.blockId"
            :data-continuation="frag.continuation ? '1' : undefined"
            v-html="fragmentHtml(frag)"
          />
        </div>
        <div class="wtp-page-number">{{ page.pageNumber }}</div>
      </div>
      <div v-if="pages.length === 0" class="wtp-page">
        <div class="wtp-content" />
        <div class="wtp-page-number">1</div>
      </div>
    </div>
    <!-- 测量容器挂载点：必须真实参与布局，否则量不到行盒 -->
    <div ref="host" class="wtp-measure-root" aria-hidden="true" />
  </div>
</template>

<style scoped>
.wtp-root {
  display: block;
}
.wtp-measure-root {
  position: absolute;
  width: 0;
  height: 0;
  overflow: hidden;
  visibility: hidden;
}
</style>
