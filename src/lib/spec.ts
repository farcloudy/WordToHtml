/**
 * 页面与样式的唯一真相源。
 *
 * docx 导出和预览 CSS 都从这份规格派生，保证「预览所见 = 导出所得」。
 * 所有取值都可以通过组件 props 局部覆盖，没有任何一处是硬编码的。
 *
 * 单位约定：
 * - 几何尺寸（纸张、页边距）用 mm
 * - 排版（字号、行距、段距）用 pt
 *
 * 换算关系（换算函数见下方，不要在别处重写）：
 * - docx 的 spacing / indent 走 twips：1pt = 20 twips
 * - docx 的 run.size 走半磅：14pt = 28
 * - CSS 走 px：1pt = 96/72 px，1mm = 96/25.4 px
 */

export type BlockKind =
  | 'title'
  | 'h1'
  | 'h2'
  | 'h3'
  | 'body'
  | 'salutation'
  | 'signature'
  | 'listTitle'
  | 'listItem'

export const BLOCK_KINDS: readonly BlockKind[] = [
  'title',
  'h1',
  'h2',
  'h3',
  'body',
  'salutation',
  'signature',
  'listTitle',
  'listItem',
]

/**
 * 有样式定义的键 = 正文块 + 页脚。
 *
 * 页脚不是「块」：它不出现在正文流里，也不会被 md 解析产出，所以不进 BlockKind。
 * 但它确实是一条要写进 docx 样式库、要在预览里生效的段落样式，所以和正文块一起
 * 放进 STYLE_KEYS —— 这样导出、预览 CSS、两个验收脚本都只需要遍历一份清单。
 */
export type StyleKey = BlockKind | 'footer'

export const STYLE_KEYS: readonly StyleKey[] = [...BLOCK_KINDS, 'footer']

export type LineRule = 'exact' | 'atLeast' | 'auto'
export type Align = 'left' | 'center' | 'right' | 'both'

/** 自动编号样式；'none' 表示该级不编号 */
export type NumberingStyle = 'chineseDot' | 'parenChinese' | 'arabicDot' | 'none'

export interface TextStyleSpec {
  /**
   * docx 样式名（w:name）。**Word 就是按这个字段判定「是不是内置样式」的**，
   * 而且是按本地化名匹配：写「标题 1」Word 就认成内置的 Heading 1，
   * 写「标题1」（少一个空格）就不认（实测 builtIn=false，2026-09-13）。
   * 所以这里内置的那几个必须逐字照抄 Word 的中文名，含中间的空格。
   * 没有内置对应的（抬头/落款/列表标题）才用自定义名。
   * 反向解析 docx 时以它作为锚点，改了会导致旧文件读不回来。
   */
  name: string
  /**
   * docx 样式 id（w:styleId）。不必是内置 id —— 名字已经决定了内置归属，
   * 而沿用 WT- 前缀还能避开 docx 库的坑：一旦 styleId 撞上它的内置样式表
   * （Heading1、Title 等），它会额外注入一份自己的默认定义，同一份 styles.xml
   * 里就出现两个同 id 的 w:style。
   */
  id: string
  /** 中文字体（w:rFonts/@w:eastAsia） */
  eastAsia: string
  /** 西文字体（w:rFonts/@w:ascii、@w:hAnsi） */
  ascii: string
  /** 字号，pt */
  sizePt: number
  bold: boolean
  align: Align
  /** 首行缩进，字符数；0 表示无缩进 */
  firstLineChars: number
  /** 行距规则 */
  lineRule: LineRule
  /** 行距值，pt；lineRule 为 auto 时忽略 */
  linePt: number
  /** 段前间距，行数（按本段行距换算成 pt） */
  spaceBeforeLines: number
  /** 段后间距，行数 */
  spaceAfterLines: number
  /** 自动编号规则；页脚这类不成块的样式固定填 'none' */
  numbering: NumberingStyle
}

/**
 * 带单位的长度。写成模板字面量类型而不是裸 string，是为了同时满足 docx 的
 * UniversalMeasure 约束 —— 传 '25m' 或 '25px' 这类值会在编译期就被拦下来。
 */
export type Length = `${number}${'mm' | 'cm' | 'in' | 'pt' | 'pc' | 'pi'}`

export type MarginSpec = { top: Length; right: Length; bottom: Length; left: Length }

export interface PageSpec {
  /** 纸张尺寸 */
  size: { width: Length; height: Length }
  /** 页边距。四边各自独立，不是单一值 —— 用 MARGIN_PRESETS 整体切换 */
  margin: MarginSpec
  /** 页脚底边到纸张底边的距离 */
  footer: Length
  /** 页眉顶边到纸张顶边的距离 */
  header: Length
}

/**
 * 页边距预设。四边分别取值，切换时整组替换；
 * 页码样式不在这里 —— 它是 `styles.footer`（内置「页脚」样式）。
 */
export interface MarginPreset {
  key: string
  label: string
  margin: MarginSpec
}

/** 默认页边距：四边等距 25mm */
export const DEFAULT_MARGIN: MarginSpec = {
  top: '25mm',
  right: '25mm',
  bottom: '25mm',
  left: '25mm',
}

/** 至少一项，这样调用方取 [0] 当默认值时不必处理 undefined */
export const MARGIN_PRESETS: readonly [MarginPreset, ...MarginPreset[]] = [
  { key: 'default', label: '四边 25mm', margin: DEFAULT_MARGIN },
  {
    key: 'gov',
    label: '公文标准（上37 下35 左28 右26 mm）',
    margin: { top: '37mm', right: '26mm', bottom: '35mm', left: '28mm' },
  },
]

export interface Spec {
  page: PageSpec
  styles: Record<StyleKey, TextStyleSpec>
}

/**
 * 默认规格。
 *
 * 与需求文档的对应关系：
 *   title       (1) 文本标题        → 内置「标题」
 *   body        (2) 正文            → 内置「正文」（即 Word 的 Normal，见 docx/export.ts）
 *   h1/h2/h3    (2.1)(2.2)(2.3)     → 内置「标题 1/2/3」，自动编号
 *   salutation  (2.4) 抬头（正文但无首行缩进）    → 自定义名
 *   signature   (2.5) 落款（正文但右对齐）        → 自定义名
 *   listItem    (3) 列表段落        → 内置「列表段落」
 *   listTitle   (3.1) 列表标题      → 自定义名（Word 无对应内置）
 *   footer      页脚页码段落        → 内置「页脚」
 */
export const DEFAULT_SPEC: Spec = {
  page: {
    size: { width: '210mm', height: '297mm' },
    margin: DEFAULT_MARGIN,
    footer: '12.5mm',
    header: '12.5mm',
  },
  styles: {
    title: {
      name: '标题',
      id: 'WT-Title',
      eastAsia: '华文中宋',
      ascii: 'Times New Roman',
      sizePt: 18,
      bold: true,
      align: 'center',
      firstLineChars: 0,
      lineRule: 'exact',
      linePt: 26,
      spaceBeforeLines: 1.5,
      spaceAfterLines: 1.5,
      numbering: 'none',
    },
    h1: {
      name: '标题 1',
      id: 'WT-H1',
      eastAsia: '仿宋',
      ascii: 'Times New Roman',
      sizePt: 14,
      bold: true,
      align: 'both',
      firstLineChars: 2,
      lineRule: 'exact',
      linePt: 25,
      spaceBeforeLines: 0.5,
      spaceAfterLines: 0.5,
      numbering: 'chineseDot',
    },
    h2: {
      name: '标题 2',
      id: 'WT-H2',
      eastAsia: '仿宋',
      ascii: 'Times New Roman',
      sizePt: 14,
      bold: true,
      align: 'both',
      firstLineChars: 2,
      lineRule: 'exact',
      linePt: 25,
      spaceBeforeLines: 0.5,
      spaceAfterLines: 0.5,
      numbering: 'parenChinese',
    },
    h3: {
      name: '标题 3',
      id: 'WT-H3',
      eastAsia: '仿宋',
      ascii: 'Times New Roman',
      sizePt: 14,
      bold: true,
      align: 'both',
      firstLineChars: 2,
      lineRule: 'exact',
      linePt: 25,
      spaceBeforeLines: 0.5,
      spaceAfterLines: 0.5,
      numbering: 'arabicDot',
    },
    body: {
      name: '正文',
      id: 'WT-Body',
      eastAsia: '仿宋',
      ascii: 'Times New Roman',
      sizePt: 14,
      bold: false,
      align: 'both',
      firstLineChars: 2,
      lineRule: 'exact',
      linePt: 25,
      spaceBeforeLines: 0.5,
      spaceAfterLines: 0.5,
      numbering: 'none',
    },
    salutation: {
      name: '抬头',
      id: 'WT-Salutation',
      eastAsia: '仿宋',
      ascii: 'Times New Roman',
      sizePt: 14,
      bold: false,
      align: 'left',
      firstLineChars: 0,
      lineRule: 'exact',
      linePt: 25,
      spaceBeforeLines: 0.5,
      spaceAfterLines: 0.5,
      numbering: 'none',
    },
    signature: {
      name: '落款',
      id: 'WT-Signature',
      eastAsia: '仿宋',
      ascii: 'Times New Roman',
      sizePt: 14,
      bold: false,
      align: 'right',
      firstLineChars: 0,
      lineRule: 'exact',
      linePt: 25,
      spaceBeforeLines: 0.5,
      spaceAfterLines: 0.5,
      numbering: 'none',
    },
    listTitle: {
      name: '列表标题',
      id: 'WT-ListTitle',
      eastAsia: '仿宋',
      ascii: 'Times New Roman',
      sizePt: 10.5,
      bold: true,
      align: 'center',
      firstLineChars: 0,
      lineRule: 'atLeast',
      linePt: 12,
      spaceBeforeLines: 0,
      spaceAfterLines: 0,
      numbering: 'none',
    },
    listItem: {
      name: '列表段落',
      id: 'WT-ListItem',
      eastAsia: '仿宋',
      ascii: 'Times New Roman',
      sizePt: 10.5,
      bold: false,
      align: 'both',
      firstLineChars: 2,
      lineRule: 'atLeast',
      linePt: 12,
      spaceBeforeLines: 0,
      spaceAfterLines: 0,
      numbering: 'none',
    },
    /**
     * 页脚页码段落。lineRule 为 auto（单倍行距）：正文的默认行距是固定值，
     * 页脚若继承它会撑出行高，显式写单倍才与预览的 line-height 对得上。
     */
    footer: {
      name: '页脚',
      id: 'WT-Footer',
      eastAsia: '宋体',
      ascii: 'Times New Roman',
      sizePt: 9,
      bold: false,
      align: 'center',
      firstLineChars: 0,
      lineRule: 'auto',
      linePt: 12,
      spaceBeforeLines: 0,
      spaceAfterLines: 0,
      numbering: 'none',
    },
  },
}

/**
 * 深度可选。用 `extends object` 而不是 `extends Record<string, unknown>`：
 * interface 不带隐式索引签名，后者会让 PageSpec 这类 interface 退化成「必须整块传」，
 * 局部覆盖页边距就会变成类型错误。
 */
export type DeepPartial<T> = T extends readonly (infer U)[]
  ? readonly DeepPartial<U>[]
  : T extends object
    ? { [K in keyof T]?: DeepPartial<T[K]> }
    : T

/** 把局部覆盖合并到默认规格上。注意 Record 的合并是逐 key 的浅层字段覆盖。 */
export function resolveSpec(override?: DeepPartial<Spec>): Spec {
  const p = override?.page
  const styles = Object.fromEntries(
    STYLE_KEYS.map((key) => [
      key,
      { ...DEFAULT_SPEC.styles[key], ...override?.styles?.[key] },
    ]),
  ) as Record<StyleKey, TextStyleSpec>

  return {
    page: {
      size: { ...DEFAULT_SPEC.page.size, ...p?.size },
      margin: { ...DEFAULT_SPEC.page.margin, ...p?.margin },
      footer: p?.footer ?? DEFAULT_SPEC.page.footer,
      header: p?.header ?? DEFAULT_SPEC.page.header,
    },
    styles,
  }
}

/* -------------------------------------------------------------------------- */
/* 单位换算                                                                    */
/* -------------------------------------------------------------------------- */

/**
 * 生成合法长度字面量的便捷函数。动态拼出来的长度（例如界面滑杆给的页边距）
 * 在类型上不是字面量，用它可以免去 `as Length` 强制转换。
 */
export function mm(value: number): Length {
  return `${value}mm`
}

/** pt → twips（docx 的 spacing / indent 单位） */
export function ptToTwips(pt: number): number {
  return Math.round(pt * 20)
}

/** pt → 半磅（docx 的 run.size 单位） */
export function ptToHalfPoints(pt: number): number {
  return Math.round(pt * 2)
}

const UNIT_TO_PX: Record<string, number> = {
  mm: 96 / 25.4,
  cm: 96 / 2.54,
  in: 96,
  pt: 96 / 72,
  pc: 16,
  pi: 16,
}

const LENGTH_RE = /^(-?\d*\.?\d+)\s*(mm|cm|in|pt|pc|pi)$/

/** 把 CSS 长度字符串换算成 px。浏览器里 1in 恒等于 96px，所以这个换算是确定的。 */
export function lengthToPx(value: string): number {
  const m = LENGTH_RE.exec(value.trim())
  const amount = m?.[1]
  const unit = m?.[2]
  if (amount === undefined || unit === undefined) {
    throw new Error(`无法解析长度：${value}`)
  }
  return Number.parseFloat(amount) * (UNIT_TO_PX[unit] ?? 1)
}

/** pt → px */
export function ptToPx(pt: number): number {
  return (pt * 96) / 72
}

/** 版心尺寸（px），分页器用它算每页容量 */
export function contentBoxPx(spec: Spec): { width: number; height: number } {
  return {
    width:
      lengthToPx(spec.page.size.width) -
      lengthToPx(spec.page.margin.left) -
      lengthToPx(spec.page.margin.right),
    height:
      lengthToPx(spec.page.size.height) -
      lengthToPx(spec.page.margin.top) -
      lengthToPx(spec.page.margin.bottom),
  }
}
