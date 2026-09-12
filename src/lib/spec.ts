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

export type LineRule = 'exact' | 'atLeast' | 'auto'
export type Align = 'left' | 'center' | 'right' | 'both'

/** 自动编号样式；'none' 表示该级不编号 */
export type NumberingStyle = 'chineseDot' | 'parenChinese' | 'arabicDot' | 'none'

export interface TextStyleSpec {
  /** docx 样式名（w:name）。反向解析 docx 时以它作为锚点，改了会导致旧文件读不回来 */
  name: string
  /** docx 样式 id（w:styleId），须是合法 XML 名 */
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
  numbering: NumberingStyle
}

/**
 * 带单位的长度。写成模板字面量类型而不是裸 string，是为了同时满足 docx 的
 * UniversalMeasure 约束 —— 传 '25m' 或 '25px' 这类值会在编译期就被拦下来。
 */
export type Length = `${number}${'mm' | 'cm' | 'in' | 'pt' | 'pc' | 'pi'}`

export interface PageSpec {
  /** 纸张尺寸 */
  size: { width: Length; height: Length }
  /** 页边距 */
  margin: { top: Length; right: Length; bottom: Length; left: Length }
  /** 页脚底边到纸张底边的距离 */
  footer: Length
  /** 页眉顶边到纸张顶边的距离 */
  header: Length
  /**
   * 页码样式。位置固定为页脚居中、字号固定 9pt、不允许修改，
   * 集中放在这里只是为了让预览 CSS 和 docx 导出共用同一个数值。
   */
  pageNumber: { sizePt: number; eastAsia: string; ascii: string }
}

export interface Spec {
  page: PageSpec
  styles: Record<BlockKind, TextStyleSpec>
}

/**
 * 默认规格。
 *
 * 与需求文档的对应关系：
 *   title       (1) 文本标题
 *   body        (2) 正文
 *   h1/h2/h3    (2.1)(2.2)(2.3) 三级标题，自动编号
 *   salutation  (2.4) 抬头（正文但无首行缩进）
 *   signature   (2.5) 落款（正文但右对齐）
 *   listItem    (3) 列表段落
 *   listTitle   (3.1) 列表标题
 */
export const DEFAULT_SPEC: Spec = {
  page: {
    size: { width: '210mm', height: '297mm' },
    margin: { top: '25mm', right: '25mm', bottom: '25mm', left: '25mm' },
    footer: '12.5mm',
    header: '12.5mm',
    pageNumber: { sizePt: 9, eastAsia: '宋体', ascii: 'Times New Roman' },
  },
  styles: {
    title: {
      name: '公文标题',
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
      name: '公文一级标题',
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
      name: '公文二级标题',
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
      name: '公文三级标题',
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
      name: '公文正文',
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
      name: '公文抬头',
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
      name: '公文落款',
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
      name: '公文列表标题',
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
      name: '公文列表段落',
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
    BLOCK_KINDS.map((kind) => [
      kind,
      { ...DEFAULT_SPEC.styles[kind], ...override?.styles?.[kind] },
    ]),
  ) as Record<BlockKind, TextStyleSpec>

  return {
    page: {
      size: { ...DEFAULT_SPEC.page.size, ...p?.size },
      margin: { ...DEFAULT_SPEC.page.margin, ...p?.margin },
      footer: p?.footer ?? DEFAULT_SPEC.page.footer,
      header: p?.header ?? DEFAULT_SPEC.page.header,
      pageNumber: { ...DEFAULT_SPEC.page.pageNumber, ...p?.pageNumber },
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
