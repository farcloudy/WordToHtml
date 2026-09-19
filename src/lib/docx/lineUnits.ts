/**
 * 把「段前/段后 = N 行」这件事真正写进 docx。
 *
 * docx 库的 `ISpacingProperties` 只暴露 before/after/line/lineRule（见它的
 * createSpacing 实现），没有 Word 的 `w:beforeLines` / `w:afterLines`。
 * 而 Word 区分「行」和「磅」看的正是这对属性：
 *
 *   · 有 beforeLines/afterLines → 界面显示为「行」，实际间距 = N × 文档网格行高；
 *   · 只有 before/after          → 界面显示为「磅」，是个固定绝对值。
 *
 * 所以样式表打包出来以后要再过一道：给每个 `<w:spacing>` 补上这对属性，
 * 并把 before/after 那对后备值按网格行高写好（见 spec.ts 的 lineSpacePt）。
 *
 * 这里只做纯字符串改写，不认识 docx 的类型，也不碰 zip —— 打包与解包在
 * export.ts 里做，这样这一段可以在 node 里直接单测。
 */

import { STYLE_KEYS } from '../spec'
import type { Spec } from '../spec'

/** 段前/段后，单位「行」 */
export interface LineUnit {
  before: number
  after: number
}

export interface LineUnitPlan {
  /** 正文（docDefaults）的段前/段后 */
  defaults: LineUnit
  /** 样式 id（w:styleId）→ 段前/段后 */
  byStyleId: Record<string, LineUnit>
}

/** 由规格表算出「哪些样式要写成行单位」；无网格的模板返回 `null`（见下）。正文走 docDefaults，不在这里。 */
export function lineUnitPlan(spec: Spec): LineUnitPlan | null {
  /*
   * 没有文档网格（`gridType: 'none'`）就**不写**这对属性。
   *
   * 「行」这个单位得有基准才成立：Word 在没有网格时把「1 行」算作 **12pt**（2026-09-19
   * 实测：段前写 1.5 行、去掉网格之后 Word 报的 SpaceBefore 从 23.4pt 掉到 18pt），
   * 而本规格表的换算基准是 15.6pt —— 两者无关，写了它 Word 里的实际段距就会与预览
   * 静默分家（预览看不到、导出的文件也「没错」，只是两个引擎各算各的）。
   *
   * 不写的话 Word 用 w:before/w:after 那对磅值，正好是预览用的同一个数。
   * 代价是 Word 里那几栏显示成「磅」而不是「行」—— 无网格时 Word 本来就给不出
   * 有意义的「行」。
   */
  if (spec.page.gridType === 'none') return null

  const byStyleId: Record<string, LineUnit> = {}
  for (const key of STYLE_KEYS) {
    if (key === 'body') continue
    const s = spec.styles[key]
    byStyleId[s.id] = { before: s.spaceBeforeLines, after: s.spaceAfterLines }
  }
  return {
    defaults: {
      before: spec.styles.body.spaceBeforeLines,
      after: spec.styles.body.spaceAfterLines,
    },
    byStyleId,
  }
}

/** 把「行」换算成 w:beforeLines 要的百分之一行 */
function hundredths(lines: number): number {
  return Math.round(lines * 100)
}

/** 替换 block 里第一个 <w:spacing>，补上行单位属性；已有就不动 */
function withLineUnits(block: string, unit: LineUnit): string {
  let done = false
  return block.replace(
    /<w:spacing\b([^>]*?)(\/?)>/g,
    (whole: string, attrs: string, selfClosing: string) => {
      if (done || /\sw:beforeLines=/.test(attrs)) return whole
      done = true
      return (
        `<w:spacing${attrs} w:beforeLines="${hundredths(unit.before)}"` +
        ` w:afterLines="${hundredths(unit.after)}"${selfClosing}>`
      )
    },
  )
}

/**
 * 改写 styles.xml：docDefaults（正文）与每条被规格表认领的段落样式，
 * 都补上 w:beforeLines / w:afterLines。认不出的样式（docx 自带的
 * FootnoteText 之类）原样放过；`plan` 为 `null`（无网格的模板）时一字不改。
 */
export function patchStylesXml(xml: string, plan: LineUnitPlan | null): string {
  if (!plan) return xml
  let out = xml.replace(/<w:pPrDefault\b[\s\S]*?<\/w:pPrDefault>/, (block) =>
    withLineUnits(block, plan.defaults),
  )
  out = out.replace(
    /<w:style\b[^>]*?w:styleId="([^"]+)"[^>]*>[\s\S]*?<\/w:style>/g,
    (block: string, id: string) => {
      const unit = plan.byStyleId[id]
      return unit ? withLineUnits(block, unit) : block
    },
  )
  return out
}
