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
  | 'attachment'
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
  'attachment',
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

/**
 * 行距规则。
 *
 * - `exact`   固定值：行盒恒为 `linePt`，字再大也不撑开；
 * - `atLeast` 最小值：不小于 `linePt`，字大就跟着撑开；
 * - `auto`    单倍行距：预览写 `line-height: normal`，docx 写 `w:lineRule="auto"`；
 * - `grid`    **单倍行距吸附文档网格**：Word 里一行正好占一个网格行，字大了也吸附到
 *             整数个网格行；浏览器的排版引擎没有网格吸附，预览只能画成固定值
 *             `page.gridLinePt`。两边因此仍然同源 —— Word 写单倍、预览画网格行高，
 *             表达的都是「一行 = 一个网格行」（2026-09-19 在 Word 实测：公文那套八条
 *             样式从固定值 28.95pt 改成它之后，22 行/页与页数一个都没变）。
 *
 * 所以 `lineRule: 'grid'` 的样式，其 `linePt` 由 `resolveSpec()` 强制成
 * `page.gridLinePt`：网格行高只有一个数，写两处迟早会分家。
 */
export type LineRule = 'exact' | 'atLeast' | 'auto' | 'grid'
export type Align = 'left' | 'center' | 'right' | 'both'

/**
 * 文档网格类型，对应 Word「页面设置 → 文档网格」那三档
 * （Word COM 的 `PageSetup.LayoutMode`：0 / 2 / 1 依次是下面三种）。
 */
export type GridType = 'none' | 'lines' | 'linesAndChars'

/** 自动编号样式；'none' 表示该级不编号 */
export type NumberingStyle =
  | 'chineseDot'
  | 'parenChinese'
  | 'arabicDot'
  | 'arabicPeriodSpace'
  | 'none'

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
  /**
   * 段前间距，单位「行」。
   *
   * 「行」的基准是**文档网格行高**（page.gridLinePt），不是本段的 linePt ——
   * Word 就是这么算的：设了 `w:docGrid/@w:linePitch` 之后，段前 0.5 行 = 0.5 ×
   * 网格行高，跟这一段自己的行距无关。所以这里不能用 `spaceBeforeLines * linePt`
   * 折算，那样在 Word 里显示成「磅」的固定值，行距一变间距就不跟着走。
   * 换算统一走 lineSpacePt()，不要在别处重写。
   */
  spaceBeforeLines: number
  /** 段后间距，行数；基准同 spaceBeforeLines */
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
  /** 页边距。四边各自独立，不是单一值 —— 用 DOC_TEMPLATES 整组切换 */
  margin: MarginSpec
  /** 页脚底边到纸张底边的距离 */
  footer: Length
  /** 页眉顶边到纸张顶边的距离 */
  header: Length
  /**
   * 文档网格类型。`'none'`（无网格）是 Word 的默认，也是「管理人文件」模板的取值。
   *
   * 它不只是「写不写 w:docGrid」：**没有网格时 Word 把「1 行」算作 12pt**（2026-09-19
   * 实测：段前 1.5 行从 23.4pt 掉到 18pt），与这里的换算基准 15.6pt 无关 —— 所以
   * `'none'` 的模板连 `w:beforeLines` / `w:afterLines` 也一概不写，只写磅值，
   * 否则 Word 里的实际段距会静默与预览分家（见 docx/lineUnits.ts）。
   */
  gridType: GridType
  /**
   * 文档网格的行高，pt。**全文只有一个**（docx 里它是节属性 `w:docGrid/@w:linePitch`），
   * 而段前/段后的「行」、以及 `lineRule: 'grid'` 的行高都以它为基准。
   *
   * 取 15.6pt（312 缇）：这是 Word 中文默认文档的网格（A4 默认页边距下「每页 44 行」），
   * 也是真实公文里实际生效的值 —— 据一份真实公文导出的 styles.xml 实测，Normal 的
   * 「0.5 行」写作 before="156"、Title 的「1.5 行」写作 before="468"，反推 1 行都是
   * 312 缇。按这个基准导出，段前/段后与那份公文逐字相同。
   *
   * `gridType === 'none'` 时它退化成纯换算常数（预览的段距、docx 的磅值后备都用它，
   * 两边照旧一致），不再代表 Word 眼里的「一行」。
   */
  gridLinePt: number
  /**
   * `gridType === 'linesAndChars'` 时每行排多少字（Word 的「指定行和字符网格」）。
   * 省略 = 不限制字数，按字号自然排。
   *
   * ⚠️ **这一项在预览与导出之间是有意不一致的**：156mm 版心放得下 27 个三号字，浏览器
   * 不会像 Word 那样压缩字符间距，所以**预览只能排 27 字**（2026-09-16 用户已接受）；
   * 导出侧照这个数写 `w:charSpace`，Word 打开就是真的 28 字。两边的页数仍要对得上
   * （实测同一份 demo 源码 27 字与 28 字在 Word 里都是 5 页，与预览一致）。
   */
  gridCharsPerLine?: number
}

/**
 * 页边距预设。四边分别取值，切换时整组替换；
 * 页码样式不在这里 —— 它是 `styles.footer`（内置「页脚」样式）。
 *
 * 它是 `DOC_TEMPLATES` 的投影（见下），保留这个导出只是因为它已在 lib 接口里。
 */
export interface MarginPreset {
  key: string
  label: string
  margin: MarginSpec
}

/** 默认页边距：四边等距 25mm，即「管理人文件」模板的页边距 */
export const DEFAULT_MARGIN: MarginSpec = {
  top: '25mm',
  right: '25mm',
  bottom: '25mm',
  left: '25mm',
}

/** 公文标准页边距：上 37 / 下 35 / 左 28 / 右 26 mm（「简易公文格式」模板用） */
const GOV_MARGIN: MarginSpec = {
  top: '37mm',
  right: '26mm',
  bottom: '35mm',
  left: '28mm',
}

/**
 * 「简易公文格式」的行网格高，pt。它同时也是这一套的 `page.gridLinePt`
 * —— Word 的「段后 1 行」以文档网格行高为基准，设成它，标题的「空一行」才是正好一行。
 *
 * 每页 22 行的算术：上 37 / 下 35 的版心高 = 637.795pt，÷ 22 = 28.99pt；**取 28.95pt
 * （= 579 缇）才装得下 22 行**（22 × 28.95 = 636.90pt，余 0.90pt），取 29pt 就只剩 21 行。
 *
 * 各条样式写的是 `lineRule: 'grid'`（单倍行距吸附文档网格），**不是固定值 28.95pt**：
 * Word 那侧该是什么就是什么（用户在 Word 里看到的行距是「单倍行距」），而浏览器没有网格
 * 吸附，预览由 `resolveSpec()` 把这一档的行高折算成网格行高 28.95pt —— 两边仍然是
 * 「一行 = 一个网格行」。`verify:pages` 拿预览页数与 Word 页数逐套模板对账，验的就是
 * 这条等价关系（2026-09-19 实测：改成它之后 Word 报 LineSpacingRule=单倍、每页 22 行、
 * 页数不变）。
 *
 * 这一套的网格类型是**指定行和字符网格**（`gridCharsPerLine: 28`）：字符那一维只有 Word
 * 排得出来，预览排 27 字（见 `PageSpec.gridCharsPerLine` 的说明）。
 */
const GOV_LINE_PT = 28.95

/**
 * 「简易公文格式」（`DOC_TEMPLATES` 第二套）的样式覆盖。
 *
 * 体例：正文与各级标题都是三号 = 16pt，文本标题二号 = 22pt，列表与页脚四号 = 14pt；
 * 正文与各级标题的行高一律压成网格行高，且都不留段前段后。字号取值见 issues/20260916.md。
 *
 * **每行 27 字、不是 GB/T 9704 写的 28 字，这是已知偏差而不是笔误**：版心宽 156mm =
 * 442.20pt，三号字每字 16pt，28 字要 448.00pt。Word 靠文档网格**压缩字符间距**把 28 字
 * 挤进 156mm，浏览器的排版引擎没有这个能力，实测 442.20pt 只放得下 27 字。左右边距各
 * 让 2mm（各 25mm）就能精确容下 28 字，但那样版心就不是公文标准的 156×225mm 了 ——
 * 2026-09-16 用户选择保标准版心。
 *
 * name / id 一概不覆盖：两套模板共用同一批 Word 样式名，导出与反向解析（P4）都按名字
 * 锚定，这里改了会让同一份文档在两套模板下写出两套 styles.xml。
 */
const GOV_STYLES: { [K in StyleKey]?: Partial<TextStyleSpec> } = {
  /**
   * 文本标题：二号方正小标宋简体，居中、不加粗。行高取**网格行**（`lineRule: 'grid'`）
   * —— 公文里标题也占一行网格，整篇的行基线才对得齐（22pt 字的单倍行高是 25.27pt，
   * 吸附到一个 28.95pt 的网格行里，不会被裁）。段后空一行。
   */
  title: {
    eastAsia: '方正小标宋简体',
    ascii: 'Times New Roman',
    sizePt: 22,
    bold: false,
    align: 'center',
    firstLineChars: 0,
    lineRule: 'grid',
    linePt: GOV_LINE_PT,
    spaceBeforeLines: 0,
    spaceAfterLines: 1,
    numbering: 'none',
  },
  /** 一级标题：黑体三号，不加粗、两端对齐、缩进 2 字、不留段前段后，编号「一、」 */
  h1: {
    eastAsia: '黑体',
    ascii: 'Times New Roman',
    sizePt: 16,
    bold: false,
    align: 'both',
    firstLineChars: 2,
    lineRule: 'grid',
    linePt: GOV_LINE_PT,
    spaceBeforeLines: 0,
    spaceAfterLines: 0,
    numbering: 'chineseDot',
  },
  /** 二级标题：楷体三号，其余同一级标题，编号「（一）」 */
  h2: {
    eastAsia: '楷体',
    ascii: 'Times New Roman',
    sizePt: 16,
    bold: false,
    align: 'both',
    firstLineChars: 2,
    lineRule: 'grid',
    linePt: GOV_LINE_PT,
    spaceBeforeLines: 0,
    spaceAfterLines: 0,
    numbering: 'parenChinese',
  },
  /** 三级标题：**与正文同款**（仿宋三号），只多一个自动编号「1. 」 */
  h3: {
    eastAsia: '仿宋',
    ascii: 'Times New Roman',
    sizePt: 16,
    bold: false,
    align: 'both',
    firstLineChars: 2,
    lineRule: 'grid',
    linePt: GOV_LINE_PT,
    spaceBeforeLines: 0,
    spaceAfterLines: 0,
    numbering: 'arabicPeriodSpace',
  },
  /** 正文：三号仿宋、两端对齐、缩进 2 字、行高 = 网格行高、无段前段后 */
  body: {
    eastAsia: '仿宋',
    ascii: 'Times New Roman',
    sizePt: 16,
    bold: false,
    align: 'both',
    firstLineChars: 2,
    lineRule: 'grid',
    linePt: GOV_LINE_PT,
    spaceBeforeLines: 0,
    spaceAfterLines: 0,
    numbering: 'none',
  },
  /** 抬头：随正文（三号仿宋），只取消首行缩进、改左对齐 */
  salutation: {
    eastAsia: '仿宋',
    ascii: 'Times New Roman',
    sizePt: 16,
    bold: false,
    align: 'left',
    firstLineChars: 0,
    lineRule: 'grid',
    linePt: GOV_LINE_PT,
    spaceBeforeLines: 0,
    spaceAfterLines: 0,
    numbering: 'none',
  },
  /** 落款：随正文，只取消首行缩进、改右对齐 */
  signature: {
    eastAsia: '仿宋',
    ascii: 'Times New Roman',
    sizePt: 16,
    bold: false,
    align: 'right',
    firstLineChars: 0,
    lineRule: 'grid',
    linePt: GOV_LINE_PT,
    spaceBeforeLines: 0,
    spaceAfterLines: 0,
    numbering: 'none',
  },
  /**
   * 附件：随正文的字号与行高，只按「附件」标识的体例改字体（黑体）、取消首行缩进、
   * 段前归零、段后留 1 行 —— 与「管理人文件」的逻辑一致，西文槽位同样用黑体。
   */
  attachment: {
    eastAsia: '黑体',
    ascii: '黑体',
    sizePt: 16,
    bold: false,
    align: 'both',
    firstLineChars: 0,
    lineRule: 'grid',
    linePt: GOV_LINE_PT,
    spaceBeforeLines: 0,
    spaceAfterLines: 1,
    numbering: 'none',
  },
  /** 列表标题：四号（14pt）加粗居中，其余同「管理人文件」 */
  listTitle: {
    eastAsia: '仿宋',
    ascii: 'Times New Roman',
    sizePt: 14,
    bold: true,
    align: 'center',
    firstLineChars: 0,
    lineRule: 'atLeast',
    linePt: 16,
    spaceBeforeLines: 0,
    spaceAfterLines: 0,
    numbering: 'none',
  },
  /**
   * 列表段落：四号（14pt），其余逻辑同「管理人文件」—— 行距取「最小值」，值就是该字号
   * 的单倍自然行高（14pt × 1.1406 = 15.97 ≈ 16pt，与管理人文件那条 10.5pt → 12pt 同源；
   * 仿宋/黑体/楷体的字面行距倍数都是 1.1406 = 292/256，实测字体 hhea 指标）。表格默认
   * 用它，所以「最小一行」的高度 = 16pt，「最小两行」= 32pt。
   */
  listItem: {
    eastAsia: '仿宋',
    ascii: 'Times New Roman',
    sizePt: 14,
    bold: false,
    align: 'both',
    firstLineChars: 0,
    lineRule: 'atLeast',
    linePt: 16,
    spaceBeforeLines: 0,
    spaceAfterLines: 0,
    numbering: 'none',
  },
  /** 页脚：四号（14pt）宋体居中，单倍行距（正文的固定行距会把页高撑坏） */
  footer: {
    eastAsia: '宋体',
    ascii: 'Times New Roman',
    sizePt: 14,
    bold: false,
    align: 'center',
    firstLineChars: 0,
    lineRule: 'auto',
    linePt: 12,
    spaceBeforeLines: 0,
    spaceAfterLines: 0,
    numbering: 'none',
  },
}

/**
 * 文件模板：**样式覆盖与页边距绑成一体**。
 *
 * 为什么不让人分别选「样式」和「页边距」：公文体例里页边距本身就是格式的一部分，
 * 让两者可以自由组合只会拼出既不是甲模板也不是乙模板的东西。
 *
 * 覆盖写成整份 spec 的 DeepPartial 而不是只有 page.margin，正是为了第二套的整套字体字号：
 * 「简易公文格式」的样式覆盖见 GOV_STYLES，调用方（预览 CSS 与 docx 样式都从 resolveSpec
 * 派生）一行都不用改。
 */
export interface DocTemplate {
  key: string
  label: string
  spec: DeepPartial<Spec>
}

/** 至少一项，这样调用方取 [0] 当默认值时不必处理 undefined */
export const DOC_TEMPLATES: readonly [DocTemplate, ...DocTemplate[]] = [
  {
    key: 'manager',
    label: '管理人文件',
    // 「管理人文件」= Word 的常规排版：四边 25mm、**无文档网格**（Word 默认）。
    // 页面上不设网格之后，段前/段后也不再写「行」单位（见 docx/lineUnits.ts）——
    // 那样 Word 的实际段距才与本规格表的磅值逐项相同。
    spec: { page: { margin: DEFAULT_MARGIN, gridType: 'none' } },
  },
  {
    key: 'govDoc',
    label: '简易公文格式',
    spec: {
      // 公文标准的版心 + 每页 22 行的行网格 + 每行 28 字的字符网格（后者只有 Word 排得出来）
      page: {
        margin: GOV_MARGIN,
        gridType: 'linesAndChars',
        gridLinePt: GOV_LINE_PT,
        gridCharsPerLine: 28,
      },
      styles: GOV_STYLES,
    },
  },
]

/**
 * 页边距预设。**由 DOC_TEMPLATES 派生** —— 边距是模板的一部分，另留一份字面量
 * 迟早会和模板分家（那是两份真相，改一处忘一处）。
 *
 * 不用 resolveSpec() 派生：它在文件后面，而这里要的只是页边距，浅合并默认边距
 * 与 resolveSpec 对 page.margin 的处理完全一致。
 */
export const MARGIN_PRESETS: readonly [MarginPreset, ...MarginPreset[]] = [
  marginPresetOf(DOC_TEMPLATES[0]),
  ...DOC_TEMPLATES.slice(1).map(marginPresetOf),
]

function marginPresetOf(template: DocTemplate): MarginPreset {
  return {
    key: template.key,
    label: template.label,
    margin: { ...DEFAULT_MARGIN, ...template.spec.page?.margin },
  }
}

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
 *   attachment  附件标记（正文但黑体、顶格、段后 1 行） → 自定义名
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
    // 无网格 = Word 的默认（「管理人文件」模板就是它，见 DOC_TEMPLATES）
    gridType: 'none',
    // 312 缇 = 15.6pt，Word 中文默认文档网格。无网格的模板里它只是「行」的换算常数
    // （见 PageSpec.gridLinePt 的说明）
    gridLinePt: 15.6,
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
    /**
     * 附件标记。基于正文：字号、行距、对齐都随正文，只改字体（黑体）、
     * 取消首行缩进、段前归零、段后留 1 行 —— 公文的「附件」标识要顶格起段，
     * 与后文（附件正文）之间空开一行。
     *
     * 西文槽位也用黑体：真实公文里那条「附件」样式，w:rFonts 的
     * ascii/eastAsia/hAnsi 写的都是黑体。
     */
    attachment: {
      name: '附件',
      id: 'WT-Attachment',
      eastAsia: '黑体',
      ascii: '黑体',
      sizePt: 14,
      bold: false,
      align: 'both',
      firstLineChars: 0,
      lineRule: 'exact',
      linePt: 25,
      spaceBeforeLines: 0,
      spaceAfterLines: 1,
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
      firstLineChars: 0,
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
  const page: PageSpec = {
    size: { ...DEFAULT_SPEC.page.size, ...p?.size },
    margin: { ...DEFAULT_SPEC.page.margin, ...p?.margin },
    footer: p?.footer ?? DEFAULT_SPEC.page.footer,
    header: p?.header ?? DEFAULT_SPEC.page.header,
    gridType: p?.gridType ?? DEFAULT_SPEC.page.gridType,
    gridLinePt: p?.gridLinePt ?? DEFAULT_SPEC.page.gridLinePt,
  }
  // 可选字段只在真的设了的时候落（与「默认值不落」的仓库约定一致）
  const charsPerLine = p?.gridCharsPerLine ?? DEFAULT_SPEC.page.gridCharsPerLine
  if (charsPerLine !== undefined) page.gridCharsPerLine = charsPerLine

  // `lineRule: 'grid'` 的行高只有一个来源：网格行高。调用方覆盖的 linePt 在这里被
  // 归一化掉 —— 预览画的就是网格行高，留着一个不同的 linePt 只会让「预览 = Word」
  // 在两个数之间无声地分家（导出的单倍行距根本不看它，最容易漏）。
  const styles = Object.fromEntries(
    STYLE_KEYS.map((key) => {
      const s = { ...DEFAULT_SPEC.styles[key], ...override?.styles?.[key] }
      return [key, s.lineRule === 'grid' ? { ...s, linePt: page.gridLinePt } : s]
    }),
  ) as Record<StyleKey, TextStyleSpec>

  return { page, styles }
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

/**
 * 段前/段后（行）→ pt。基准是文档网格行高，不是段落自己的 linePt。
 *
 * docx 的 `w:before`/`w:after`、预览 CSS 的 margin、量测出来的段距，三处换算
 * 都必须走这一个函数 —— 任何一处回退成 `lines * linePt` 都会让 Word 与预览
 * 在段间距上分家，而且因为显示单位不同（磅 vs 行）很难一眼看出来。
 */
export function lineSpacePt(lines: number, spec: Spec): number {
  return lines * spec.page.gridLinePt
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
