/**
 * 节的解析层：模型 + 规格表 → 逐节清单。
 *
 * 单一真相源：预览（每页几何与页码显示）、分页（逐节版心高与页码推进）、
 * docx 导出（逐节 pgSz / pgNumType / footers）、上下文工具条（开关的已解析值）
 * 全部走这里，谁都不许自己再解析一遍「继承」关系。
 *
 * 为什么要集中：`linkPrevious`（关联前节）不是「有没有页码」这种局部属性，
 * 它把当前节的显示结果挂在前一节上 —— 分散解析必然出现「工具条说有关联、
 * 预览却按独立算」这类自相矛盾。
 */

import { contentBoxPx } from './spec'
import type { PageSpec, Spec } from './spec'
import type { DocModel, PageOrientation, SectionSettings } from './types'
import { sectionCountOf } from './edit/section'

/** 三个开关 + 方向的**已解析值**（缺省已补齐、首节的 linkPrevious 已强制为 false） */
export interface ResolvedSectionSettings {
  pageNumbers: boolean
  linkPrevious: boolean
  restartAtOne: boolean
  orientation: PageOrientation
}

export interface ResolvedSection {
  index: number
  settings: ResolvedSectionSettings
  /** 页面规格，**已按 orientation 换算过 size**（landscape 时宽高互换） */
  page: PageSpec
  /** 版心 px（与 page 同一方向） */
  content: { width: number; height: number }
  /** 已解析「继承」后的最终显示结果：这一节到底显不显示页码 */
  showPageNumber: boolean
}

/**
 * 分页器要的逐节运行时参数。
 *
 * 分页器**只搬运、不解析**：显示不显示页码、要不要从 1 重排，都是上面解析好的结果。
 * contentHeight 是逐节的 —— 横竖混排时每节版心高不同，装箱必须按当前节算。
 */
export interface SectionRuntime {
  contentHeight: number
  showPageNumber: boolean
  restartAtOne: boolean
}

/**
 * 解析一节。
 *
 * `isFirst` 时 linkPrevious 恒为 false：首节没有前节可关联，让它在模型里带
 * 「关联前节」只会产生一句永远不成立的话（UI 也据此置灰）。
 */
export function resolveSectionSettings(
  raw: SectionSettings | undefined,
  isFirst: boolean,
): ResolvedSectionSettings {
  return {
    pageNumbers: raw?.pageNumbers ?? true,
    linkPrevious: isFirst ? false : (raw?.linkPrevious ?? true),
    restartAtOne: raw?.restartAtOne ?? false,
    orientation: raw?.orientation ?? 'portrait',
  }
}

/** 按方向换算页面规格：landscape 时宽高互换，其余字段（页边距等）不动 */
function orientedPage(page: PageSpec, orientation: PageOrientation): PageSpec {
  if (orientation !== 'landscape') return page
  return { ...page, size: { width: page.size.height, height: page.size.width } }
}

/**
 * 逐节清单：下标 = 节号（0 = 首节），长度恒等于「分节符数 + 1」。
 *
 * `doc.sections` 缺失 / 偏短一律按全默认兜底 —— 手搓模型、旧 md、旧 docx 读回
 * 三种来源都可能是这样，容错放在这一处，调用方不必各自防御。
 */
export function resolveSections(doc: DocModel, spec: Spec): ResolvedSection[] {
  const count = sectionCountOf(doc)
  const out: ResolvedSection[] = []
  // 首节没有前节，它的显示结果只看自己 —— 这也是下面这串继承的起点
  let previousShow: boolean = true

  for (let index = 0; index < count; index += 1) {
    const settings = resolveSectionSettings(doc.sections?.[index], index === 0)
    const showPageNumber: boolean = settings.linkPrevious ? previousShow : settings.pageNumbers
    previousShow = showPageNumber
    const page = orientedPage(spec.page, settings.orientation)
    out.push({ index, settings, page, content: contentBoxPx({ ...spec, page }), showPageNumber })
  }
  return out
}
