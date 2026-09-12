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
import { sliceInlines, commentScopes } from '../lib/types'
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
const root = ref<HTMLElement | null>(null)
const pages = shallowRef<PageLayout[]>([])
/** 最近一次分页实际用到的量测值，排错时用来和渲染结果对账 */
const measured = shallowRef<MeasuredItem[]>([])
/** 侧栏里正在查看的批注；只影响高亮，不触发重排 */
const activeCommentId = ref<number | null>(null)

/** 每条批注的锚定文字（Word 叫 scope）。侧栏要显示「批的哪句话」就得靠它 */
const scopes = computed(() => commentScopes(parsed.value))

/** 侧栏内容：批注本身来自模型，不是从 DOM 反推的 */
const comments = computed(() =>
  parsed.value.comments.map((c) => ({
    id: c.id,
    author: c.author,
    date: c.date,
    text: c.text,
    scope: scopes.value.get(c.id) ?? '',
    resolved: c.resolved === true,
  })),
)

/** 点侧栏 → 高亮正文锚点并滚动到它 */
async function focusComment(id: number): Promise<void> {
  activeCommentId.value = activeCommentId.value === id ? null : id
  await nextTick()
  root.value
    ?.querySelector(`.wtp-comment[data-comment="${id}"]`)
    ?.scrollIntoView({ block: 'center', behavior: 'smooth' })
}

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
  <div ref="root" class="wtp-root">
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
        <!-- wtp-footer 提供排版（= 内置「页脚」样式），wtp-page-number 只负责定位 -->
        <div class="wtp-page-number wtp-footer">{{ page.pageNumber }}</div>
      </div>
      <div v-if="pages.length === 0" class="wtp-page">
        <div class="wtp-content" />
        <div class="wtp-page-number wtp-footer">1</div>
      </div>
    </div>

    <!-- 审阅侧栏：文档里没有批注时整块不占位 -->
    <aside v-if="comments.length > 0" class="wtp-comments">
      <div class="wtp-comments-title">批注 {{ comments.length }}</div>
      <ul>
        <li v-for="c in comments" :key="c.id">
          <button
            type="button"
            :class="{ 'is-active': c.id === activeCommentId }"
            @click="focusComment(c.id)"
          >
            <span class="wtp-comment-meta">
              #{{ c.id + 1 }} {{ c.author }}<template v-if="formatDate(c.date)">
                · {{ formatDate(c.date) }}</template
              >
              <em v-if="c.resolved">已解决</em>
            </span>
            <span class="wtp-comment-scope">「{{ c.scope }}」</span>
            <span class="wtp-comment-text">{{ c.text }}</span>
          </button>
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
.wtp-comments button {
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
.wtp-comments button:hover {
  background: #f8f9fb;
}
.wtp-comments button.is-active {
  background: rgba(255, 213, 0, 0.35);
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
