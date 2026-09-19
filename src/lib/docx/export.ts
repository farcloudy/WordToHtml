/**
 * 文档模型 → docx。
 *
 * 全部走段落样式（名字与 w:styleId 见 spec.ts），而不是给每个段落直接刷格式。
 * 这样有三个好处：
 *   1. 在 Word 里改样式即可整篇生效，符合公文排版的工作习惯；
 *   2. 反向解析时只要认样式名就能还原层级（第 6 条的「样式名匹配」就挂在这里）；
 *   3. 标题/正文/列表段落/页脚直接用 Word 的内置样式名，样式库里就是用户熟悉的
 *      那几条（标题、标题 1、正文……），而不是一堆「公文××」自造样式。
 *
 * 「内置」的判定权在 Word 手里，依据是 w:name 是否等于它的本地化名 ——
 * 不是 w:styleId。所以这里刻意保留 WT- 前缀的 id，只把 name 换成内置名。
 */

import {
  AlignmentType,
  BorderStyle,
  CommentRangeEnd,
  CommentRangeStart,
  CommentReference,
  DeletedTextRun,
  Document,
  DocumentGridType,
  Footer,
  HeightRule,
  InsertedTextRun,
  LineRuleType,
  PageBreak,
  PageNumber,
  PageOrientation,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  UnderlineType,
  VerticalAlignTable,
  WidthType,
} from 'docx'
import type {
  FileChild,
  IBorderOptions,
  ICommentOptions,
  IParagraphStyleOptions,
  ISectionOptions,
  ParagraphChild,
} from 'docx'
import JSZip from 'jszip'

import { STYLE_KEYS, lengthToPx, lineSpacePt, ptToHalfPoints, ptToTwips } from '../spec'
import type { Align, GridType, LineRule, Spec, TextStyleSpec } from '../spec'
import type {
  Block,
  CellVerticalAlign,
  DocModel,
  Inline,
  PageBreakBlock,
  TableBlock,
  TableCellModel,
  TableRowModel,
  TableRowRole,
  TextBlock,
} from '../types'
import { defaultCellAlignH, headerRowCount } from '../types'
import { cellParagraphs, emptyCell } from '../edit/table'
import { computeNumbering } from '../numbering'
import { resolveSections } from '../section'
import { lineUnitPlan, patchStylesXml } from './lineUnits'

export interface ExportMeta {
  title?: string
  creator?: string
  description?: string
}

function fontOf(s: TextStyleSpec): {
  ascii: string
  hAnsi: string
  eastAsia: string
  cs: string
} {
  // Word 用 ascii/hAnsi 管西文、eastAsia 管中日韩，正好对应需求里的
  // 「默认字体：仿宋 / Times New Roman」这一对。
  return { ascii: s.ascii, hAnsi: s.ascii, eastAsia: s.eastAsia, cs: s.ascii }
}

function alignmentOf(a: Align): (typeof AlignmentType)[keyof typeof AlignmentType] {
  switch (a) {
    case 'left':
      return AlignmentType.LEFT
    case 'center':
      return AlignmentType.CENTER
    case 'right':
      return AlignmentType.RIGHT
    case 'both':
      return AlignmentType.JUSTIFIED
  }
}

function lineRuleOf(r: LineRule): (typeof LineRuleType)[keyof typeof LineRuleType] {
  switch (r) {
    case 'exact':
      return LineRuleType.EXACT
    case 'atLeast':
      return LineRuleType.AT_LEAST
    case 'auto':
    case 'grid':
      // 'grid' = 单倍行距吸附文档网格：Word 那侧就是「单倍行距」（每行吸附到一个网格行），
      // 「一行 = 一个网格行」这件事由 w:docGrid 的 linePitch 兜着，不需要在这里写磅值。
      // 预览那侧没有网格吸附，由 resolveSpec() 把它的 linePt 折算成网格行高（见 spec.ts）。
      return LineRuleType.AUTO
  }
}

/**
 * 文档网格类型。'none'（无网格）不写 `w:type` —— type 的默认值就是 default（无网格），
 * 这也正是 Word 自己把一份有网格的文档改成「无网格」之后写出来的样子
 * （2026-09-19 实测：Word 写出 `<w:docGrid w:linePitch="579"/>`，读回 LayoutMode=0）。
 */
function gridTypeOf(t: GridType): (typeof DocumentGridType)[keyof typeof DocumentGridType] | undefined {
  switch (t) {
    case 'none':
      return undefined
    case 'lines':
      return DocumentGridType.LINES
    case 'linesAndChars':
      return DocumentGridType.LINES_AND_CHARS
  }
}

/**
 * 字距向内收的余量，缇。
 *
 * 「正好排满」的字距（版心宽 ÷ 每行字数）恰好落在窗口**上缘**上，而窗口是
 * `(版心宽 ÷ (N+1), 版心宽 ÷ N]` —— 取上缘就意味着「Word 侧的版心宽不能比我们的算术小
 * 哪怕一缇」。实测 Word 把 `'210mm'` 这类长度换算成缇时比我们的算术**宽 2 缇**
 * （210mm 准确是 11905.5 缇、Word 用了 11907；28mm 取 1587 也是截断），所以本机取上缘
 * 也能排成 N 字（独立验收实测：charSpace = −848 时 Word 报 CharsLine = 28）。
 * 但这个偏差的方向与大小是 Word 的实现细节，换个版本可能就反过来 —— 收 2 缇之后
 * 版心宽有 ±(2N) 缇（公文那套 0.5mm）的容差，怎么取整都不会掉成 N−1 字。
 */
const CHAR_PITCH_MARGIN_TWIPS = 2

/**
 * 字符网格：把「每行 N 字」换算成 Word 的 `w:charSpace`（`gridCharsPerLine` 没设则返回
 * undefined = 不限制字数）。算式是在本机 Word 16.0 上**反推**出来的（2026-09-19）：
 * 把该版面的 CharsLine 依次设成 20/24/26/27/28/29/30，读出 Word 自己写的那七个
 * charSpace，全落在同一条直线上：
 *
 *     charSpace = 204.8 × (字距 − 默认字号)     （三个量都是缇，1pt = 20 缇）
 *     字距      = 版心宽 ÷ 每行字数
 *
 * 204.8 = 4096 ÷ 20，所以 charSpace 的单位是 1/4096 磅；默认字号取 docDefaults 的字号
 * （就是正文的字号，见 buildDocument 的 `default.document.run.size`）—— Word 的字符网格
 * 正是以文档默认字号为一个字符格的。
 *
 * 代价：字距比「正好排满」小 2 缇，左对齐的整行末尾会比版心右缘短 2N 缇（公文那套
 * ≈1mm）。两端对齐的段落由 Word 撑满，最后一行的长短本来就无所谓，实际看不出来。
 */
function charSpaceOf(spec: Spec): number | undefined {
  const perLine = spec.page.gridCharsPerLine
  if (perLine === undefined) return undefined
  // 版心宽按**纵向**算：字符网格是页面设置的一部分，不随某一节的横排而变化
  // （横排节的版心更宽，同样的字距自然就排下更多字）。
  const pitch = contentWidthTwips(spec) / perLine - CHAR_PITCH_MARGIN_TWIPS
  return Math.round(204.8 * (pitch - ptToTwips(spec.styles.body.sizePt)))
}

/**
 * 段前/段后按「行」换算成 twips。
 *
 * 基准是**文档网格行高**而不是本段行距（见 spec.ts 的 lineSpacePt）。这里的
 * before/after 是后备值：Word 认 beforeLines/afterLines（由 lineUnits.ts 补写），
 * 不支持这对属性的渲染器才退回到这里的磅值 —— 两者必须同源，否则「谁在用什么」
 * 会变成一个看不出来的分叉。
 *
 * 行距一律显式写出（含 auto / grid → 单倍 240 twips）：正文的默认行距可能不是单倍，
 * 样式若不写就会继承它 —— 页脚那类要单倍行距的样式会被撑高。
 */
function spacingOf(
  s: TextStyleSpec,
  spec: Spec,
): {
  before: number
  after: number
  line: number
  lineRule: (typeof LineRuleType)[keyof typeof LineRuleType]
} {
  return {
    before: ptToTwips(lineSpacePt(s.spaceBeforeLines, spec)),
    after: ptToTwips(lineSpacePt(s.spaceAfterLines, spec)),
    line: s.lineRule === 'exact' || s.lineRule === 'atLeast' ? ptToTwips(s.linePt) : 240,
    lineRule: lineRuleOf(s.lineRule),
  }
}

/**
 * 导出全部段落样式定义。反向解析 docx 时按 name 反查 BlockKind。
 *
 * 不含 `body`：正文就是 Word 的 Normal，由 buildDocument 里的默认样式承担，
 * 再单独定义一条「正文」段落样式只会多出一个同名的 Normal。
 */
export function paragraphStyles(spec: Spec): IParagraphStyleOptions[] {
  return STYLE_KEYS.filter((key) => key !== 'body').map((key) => {
    const s = spec.styles[key]
    return {
      id: s.id,
      name: s.name,
      basedOn: 'Normal',
      quickFormat: true,
      run: {
        font: fontOf(s),
        size: ptToHalfPoints(s.sizePt),
        bold: s.bold,
      },
      paragraph: {
        alignment: alignmentOf(s.align),
        spacing: spacingOf(s, spec),
        // 必须显式写，哪怕为 0：正文的默认首行缩进会被基于它的样式继承，
        // 抬头/落款/页脚这类不该有缩进的样式会被静默缩进 2 字符。
        indent: { firstLineChars: s.firstLineChars * 100 },
      },
    }
  })
}

/**
 * 页码段落。字体字号对齐全部来自「页脚」样式（Word 内置样式名）。
 *
 * `showNumbers` 为 false 时给一个**空段落**：这一节不显示页码时不能什么都不写 ——
 * 不写页脚引用就等于「沿用前一节的页脚」，那前一节的页码会带着跑过来，
 * 「无页码」这个要求就落空了（见 PLAN 7.2 指出的那个坑）。
 */
function pageNumberParagraph(spec: Spec, showNumbers: boolean): Paragraph {
  return new Paragraph({
    style: spec.styles.footer.id,
    ...(showNumbers ? { children: [new TextRun({ children: [PageNumber.CURRENT] })] } : {}),
  })
}

/**
 * 行内模型 → docx 子元素。普通段落与单元格共用这一个（不要各抄一份）：
 * 批注锚点、加粗、下划线、颜色、修订在表格内外必须是同一套语义。
 */
function inlineChildren(inlines: readonly Inline[], prefix = ''): ParagraphChild[] {
  const children: ParagraphChild[] = []
  if (prefix) {
    // 编号按需求写成正文文字，与预览完全一致
    children.push(new TextRun({ text: prefix }))
  }

  for (const inline of inlines) {
    if (inline.t === 'break') {
      // 软换行：lib 的 TextRun 有 break 字段（见 index.d.ts 的 IRunOptionsBase），
      // 它会在 <w:r> 里放一个不带属性的 <w:br/>，正是 Word 的软换行（不是分页的 w:br type="page"）。
      children.push(new TextRun({ break: 1 }))
      continue
    }
    if (inline.t === 'commentStart') {
      children.push(new CommentRangeStart(inline.commentId))
      continue
    }
    if (inline.t === 'commentEnd') {
      children.push(new CommentRangeEnd(inline.commentId))
      children.push(new CommentReference(inline.commentId))
      continue
    }

    const base = {
      text: inline.text,
      ...(inline.bold ? { bold: true } : {}),
      // 下划线与加粗同为一处 run 属性：docx 的 underline 是对象，只有写了 type
      // 才会输出 <w:u>（不写等于不加下划线，没有「true」这种简写）
      ...(inline.underline ? { underline: { type: UnderlineType.SINGLE } } : {}),
      ...(inline.color ? { color: inline.color } : {}),
    }

    if (inline.rev) {
      const changed = {
        id: inline.rev.id,
        author: inline.rev.author,
        date: inline.rev.date,
      }
      children.push(
        inline.rev.kind === 'ins'
          ? new InsertedTextRun({ ...base, ...changed })
          : new DeletedTextRun({ ...base, ...changed }),
      )
    } else {
      children.push(new TextRun(base))
    }
  }
  return children
}

function textBlockParagraph(
  block: TextBlock,
  spec: Spec,
  numbering: Map<string, string>,
  pageBreakBefore = false,
): Paragraph {
  const children = inlineChildren(block.inlines, numbering.get(block.id) ?? '')

  // 正文不挂样式：它就是 Word 的 Normal（内置「正文」），
  // 格式定义在 buildDocument 的 styles.default.document 上。
  const styleId = block.kind === 'body' ? undefined : spec.styles[block.kind].id
  return new Paragraph({
    ...(styleId ? { style: styleId } : {}),
    ...(pageBreakBefore ? { pageBreakBefore: true } : {}),
    children,
  })
}

/* -------------------------------------------------------------------------- */
/* 表格                                                                        */
/* -------------------------------------------------------------------------- */

/**
 * 单元格左右内边距：Word 默认的 108 缇（0.19cm / 5.4pt），上下 0。
 * 预览侧（W4a-2）必须取同一个值，否则格内文字会相对 Word 偏移；
 * 这里显式写出、不吃库的默认值，就是为了让两侧有唯一一个数可对。
 */
const CELL_MARGIN_TWIPS = 108

/** 正文格的框：单线 0.5pt（w:sz=4） */
const BODY_BORDER: IBorderOptions = { style: BorderStyle.SINGLE, size: 4, color: 'auto' }
/** 整行合并的 unit / note 行：四边无框 */
const NO_BORDER: IBorderOptions = { style: BorderStyle.NONE, size: 0, color: 'auto' }
const NO_BORDERS = {
  top: NO_BORDER,
  left: NO_BORDER,
  bottom: NO_BORDER,
  right: NO_BORDER,
}
const BODY_BORDERS = {
  top: BODY_BORDER,
  left: BODY_BORDER,
  bottom: BODY_BORDER,
  right: BODY_BORDER,
}

/**
 * 版心宽（twips）= 页面宽 − 左右页边距。表格总宽与字符网格的字距都从这里取。
 * 1in = 1440twips = 96px，所以 px * 15（四舍五入）就是缇；换算只在这里做一次。
 */
function contentWidthTwips(spec: Spec): number {
  const px =
    lengthToPx(spec.page.size.width) -
    lengthToPx(spec.page.margin.left) -
    lengthToPx(spec.page.margin.right)
  return Math.round(px * 15)
}

/** 等分列宽；除不尽的余数给最后一列，保证总和恰好等于版心宽 */
function columnWidthsTwips(total: number, columns: number): number[] {
  const base = Math.floor(total / columns)
  const widths = new Array<number>(columns).fill(base)
  widths[columns - 1] = (widths[columns - 1] ?? 0) + (total - base * columns)
  return widths
}

/**
 * 一个格子 → 它里面的**每一段**一个 `Paragraph`（同一个 `w:tc` 里按顺序放 N 个）。
 *
 * 格内多段落（Enter 在格内新起一段）在 Word 里的样子就是同一个 `w:tc` 里有多个 `w:p`；
 * 段落样式与水平对齐是**格子级**的（`TableCellModel.kind` / `align`），所以每一段共用同一套。
 * `kind === 'body'` 时**不挂样式**：正文就是 Word 的 Normal，styles.xml 里没有
 * WT-Body 这条（paragraphStyles 刻意不定义它），挂了就是一条指向不存在样式的
 * 悬空引用 —— 与 textBlockParagraph 对正文的处理保持一致。
 *
 * 坏输入（格子没有 paragraphs 字段）由 cellParagraphs 兜底成一段空段：`w:tc` 里必须至少有一个
 * `w:p`，否则这份 docx Word 打不开。
 */
function cellParagraph(
  cell: TableCellModel,
  spec: Spec,
  alignment?: (typeof AlignmentType)[keyof typeof AlignmentType],
): Paragraph[] {
  const kind = cell.kind ?? 'listItem'
  return cellParagraphs(cell).map(
    (para) =>
      new Paragraph({
        ...(kind === 'body' ? {} : { style: spec.styles[kind].id }),
        ...(alignment ? { alignment } : {}),
        children: inlineChildren(para.inlines),
      }),
  )
}

/** 垂直对齐三档 → docx 的枚举（缺省值一律 top，与预览侧同源） */
const VERTICAL_ALIGN: Record<CellVerticalAlign, (typeof VerticalAlignTable)[keyof typeof VerticalAlignTable]> = {
  top: VerticalAlignTable.TOP,
  middle: VerticalAlignTable.CENTER,
  bottom: VerticalAlignTable.BOTTOM,
}

/** 逐格的实际水平对齐：有覆盖用覆盖，否则 unit 右 / note 左 / body 跟该格样式 */
function cellAlignH(role: TableRowRole, cell: TableCellModel, spec: Spec): Align {
  return cell.align?.h ?? defaultCellAlignH(role, spec.styles[cell.kind ?? 'listItem'].align)
}

/**
 * 一行的最小行高基准（磅）= 该行各格样式的最大 linePt。
 *
 * 必须按**渲染后的 0..columns-1** 走：缺格补空的那些格在 Word 里也是实打实的一格，
 * 它们的样式是 listItem（预览侧也一样），不数进来两侧的最大值就会分家。
 * unit / note 行整行一格，只看第 0 格。
 */
function rowLinePt(row: TableRowModel, columns: number, spec: Spec): number {
  if (row.role !== 'body') {
    return spec.styles[row.cells[0]?.kind ?? 'listItem'].linePt
  }
  let max = 0
  for (let c = 0; c < columns; c += 1) {
    max = Math.max(max, spec.styles[row.cells[c]?.kind ?? 'listItem'].linePt)
  }
  return max
}

/**
 * 表格块 → docx 表格。
 *
 * 边框刻意分两层：**表格级不画框，正文格各自画四边**。
 * 反过来的写法（表格级画框、unit/note 格覆盖成 none）在 Word 里压不住共享边 ——
 * 一条边归相邻两格共有，只有一侧写 none、另一侧缺省时 Word 仍会把框画出来
 *（2026-09-13 用 Word COM 实测）。表格级不画就不存在这种「一侧压不住」的边，
 * unit/note 行的四边（含与正文行相邻的那边）自然全无框。
 */
function tableBlock(block: TableBlock, spec: Spec): Table {
  const total = contentWidthTwips(spec)
  // 前 N 行挂 w:tblHeader（Word 的「跨页重复标题行」）；N 由 headerRowCount 现夹，
  // 与预览 / 分页侧同一个口径
  const header = headerRowCount(block)

  const rows = block.rows.map((row, r) => {
    // 行高按该行各格样式的最大 linePt（缺格按 listItem），与预览侧同一条规则；
    // 表头行 / 附注行恒**一行**（W7）—— `minLines` 只管正文行：这两行是整张表的装饰，
    // 跟着行高设置一起变高只会白白把表格撑长（预览侧同一条规则，见 render/css.ts 的表格一段）。
    const rowHeight = ptToTwips(
      (row.role === 'body' ? block.minLines : 1) * rowLinePt(row, block.columns, spec),
    )
    let cells: TableCell[]
    if (row.role === 'body') {
      cells = []
      for (let c = 0; c < block.columns; c += 1) {
        // 缺格补空：Word 的表格必须是矩形，补齐只发生在导出这一侧，
        // 模型仍按 md 原样存（少一格的书写方式不该被解析改写）。
        const cell = row.cells[c] ?? emptyCell()
        cells.push(
          new TableCell({
            borders: BODY_BORDERS,
            verticalAlign: VERTICAL_ALIGN[cell.align?.v ?? 'top'],
            children: cellParagraph(cell, spec, alignmentOf(cellAlignH('body', cell, spec))),
          }),
        )
      }
    } else {
      // unit / note 天然整行一格 → 展开成 columnSpan = columns
      const cell = row.cells[0] ?? emptyCell()
      cells = [
        new TableCell({
          columnSpan: block.columns,
          borders: NO_BORDERS,
          verticalAlign: VERTICAL_ALIGN[cell.align?.v ?? 'top'],
          children: cellParagraph(cell, spec, alignmentOf(cellAlignH(row.role, cell, spec))),
        }),
      ]
    }

    return new TableRow({
      // 前 header 行是标题行（w:tblHeader）：Word 在每一个续页顶端自动重复它们。
      // 其余（w:trHeight ATLEAST、cantSplit、w:tblW、边框）一字不改。
      ...(r < header ? { tableHeader: true } : {}),
      cantSplit: block.cantSplit,
      height: { value: rowHeight, rule: HeightRule.ATLEAST },
      children: cells,
    })
  })

  return new Table({
    rows,
    // width.type 必须显式写 DXA：库的默认是 AUTO，不写就落到 w:tblW w:type="auto"
    width: { size: total, type: WidthType.DXA },
    columnWidths: columnWidthsTwips(total, block.columns),
    layout: TableLayoutType.FIXED,
    margins: { top: 0, bottom: 0, left: CELL_MARGIN_TWIPS, right: CELL_MARGIN_TWIPS },
    borders: {
      top: NO_BORDER,
      left: NO_BORDER,
      bottom: NO_BORDER,
      right: NO_BORDER,
      insideHorizontal: NO_BORDER,
      insideVertical: NO_BORDER,
    },
  })
}

interface SectionGroup {
  /** 本节里的块，含分页符（分节符本身不进组，它只负责切组） */
  blocks: (TextBlock | PageBreakBlock | TableBlock)[]
}

/** 按分节符把内容切成若干节。分节符在 Word 里意味着新起一页 + 独立的页码序列。 */
function groupSections(doc: DocModel): SectionGroup[] {
  const groups: SectionGroup[] = [{ blocks: [] }]
  for (const block of doc.blocks) {
    if (block.t === 'sectionBreak') {
      groups.push({ blocks: [] })
      continue
    }
    groups[groups.length - 1]?.blocks.push(block)
  }
  return groups
}

/**
 * 把一节里的块变成段落 / 表格。
 *
 * 分页符优先写成「段前分页」挂在它后面那一段上（`w:pageBreakBefore`）——
 * 这样不会像插一个空段落那样在页顶多留一个空行。
 *
 * 后面没有段落可挂时（紧跟分节符，或者干脆在节末/文末）必须退回到独立段落里
 * 写一个 `w:br w:type="page"`：**换页这件事不能丢**。否则「分页符 + 分节符」
 * 连在一起时，只剩分节符的换页生效，看上去就是「只分了一次页」。
 *
 * `Table` 上没有 `pageBreakBefore` 这种字段，所以分页符后面紧跟表格时也只能
 * 走上面那条退路，否则换页会被静默丢掉。
 */
function sectionParagraphs(
  group: SectionGroup,
  spec: Spec,
  numbering: Map<string, string>,
): FileChild[] {
  const children: FileChild[] = []
  let breakBefore = false
  for (const block of group.blocks) {
    if (block.t === 'pageBreak') {
      breakBefore = true
      continue
    }
    if (block.t === 'table') {
      // Table 上没有 pageBreakBefore，分页符只能退回到独立段落里写 <w:br type="page">
      if (breakBefore) children.push(new Paragraph({ children: [new PageBreak()] }))
      breakBefore = false
      children.push(tableBlock(block, spec))
      continue
    }
    children.push(textBlockParagraph(block, spec, numbering, breakBefore))
    breakBefore = false
  }
  if (breakBefore) children.push(new Paragraph({ children: [new PageBreak()] }))
  return children
}

export function buildDocument(
  doc: DocModel,
  spec: Spec,
  meta: ExportMeta = {},
): Document {
  const numbering = computeNumbering(doc.blocks, (b: Block) =>
    b.t === 'textBlock' ? spec.styles[b.kind].numbering : 'none',
  )
  const liveSections = resolveSections(doc, spec)
  // 字符网格与具体哪一节无关（字距按纵向版心算），先算一次
  const charSpace = charSpaceOf(spec)
  const sections: ISectionOptions[] = groupSections(doc).map((group, index) => {
    const children = sectionParagraphs(group, spec, numbering)
    if (children.length === 0) children.push(new Paragraph({}))

    // groupSections 与 resolveSections 数的是同一件事（分节符数 + 1），下标一一对应；
    // 兜底到首节只是防御（真到那一步说明两处对分节符的看法分了家）
    const section = liveSections[index] ?? liveSections[0]
    const settings = section?.settings
    /** 关联前节 = 不写页脚引用，Word 自会沿用前一节的页脚（PLAN 7.2） */
    const linkPrevious = settings?.linkPrevious ?? true

    return {
      properties: {
        page: {
          size: {
            // 宽高按**纵向**给：docx 库的 createPageSize 在 orientation=landscape 时
            // 会自己把 w/h 互换（node_modules/docx 实测，2026-09-13）。
            // 这里若再手动换一次，就会换两遍 → 纸变成「纵向尺寸 + 横向标记」。
            width: spec.page.size.width,
            height: spec.page.size.height,
            orientation:
              settings?.orientation === 'landscape'
                ? PageOrientation.LANDSCAPE
                : PageOrientation.PORTRAIT,
          },
          margin: {
            top: spec.page.margin.top,
            right: spec.page.margin.right,
            bottom: spec.page.margin.bottom,
            left: spec.page.margin.left,
            header: spec.page.header,
            footer: spec.page.footer,
          },
          // w:pgNumType 只在「从 1 重排」时写；不写就接着上一节往下数
          ...(settings?.restartAtOne ? { pageNumbers: { start: 1 } } : {}),
        },
        // 文档网格。两种作用：
        //   ① Word 的「行」单位段距（w:beforeLines）以它的 linePitch 为基准，没有它
        //      Word 会按一套我们控制不了的行高去算，段间距就对不上了；
        //   ② lineRule: 'grid' 的段落靠它把每一行吸附成一个网格行（预览那一侧由
        //      resolveSpec() 折算成同样的行高），linesAndChars 时还决定每行几个字。
        // 'none'（无网格）时只留 linePitch、不写 type —— 见 gridTypeOf 的说明。
        grid: {
          type: gridTypeOf(spec.page.gridType),
          linePitch: ptToTwips(spec.page.gridLinePt),
          ...(charSpace !== undefined ? { charSpace } : {}),
        },
      },
      // 关联前节时**整个 footers 都不写** —— 不写才没有 <w:footerReference>，
      // Word 才会沿用前一节的页脚。写了（哪怕是空页脚）就等于本节自带页脚。
      ...(linkPrevious
        ? {}
        : {
            footers: {
              default: new Footer({
                children: [pageNumberParagraph(spec, settings?.pageNumbers ?? true)],
              }),
            },
          }),
      children,
    }
  })

  return new Document({
    title: meta.title ?? '未命名文档',
    creator: meta.creator ?? 'WordToHtml',
    description: meta.description ?? '',
    styles: {
      // 正文 = Word 的 Normal。这里定义的默认样式就是样式库里的「正文」，
      // 正文段落不挂任何样式即落到它上面。
      default: {
        document: {
          run: {
            font: fontOf(spec.styles.body),
            size: ptToHalfPoints(spec.styles.body.sizePt),
            bold: spec.styles.body.bold,
          },
          paragraph: {
            alignment: alignmentOf(spec.styles.body.align),
            spacing: spacingOf(spec.styles.body, spec),
            indent: { firstLineChars: spec.styles.body.firstLineChars * 100 },
          },
        },
      },
      paragraphStyles: paragraphStyles(spec),
    },
    // 注意：这里必须传纯选项对象，不能传 new Comment(...) 实例。
    // docx 9.7.1 的 Comments 构造器会对每个 child 再包一层 new Comment(child)，
    // 传实例会把实例本身当选项对象解构，children 变成 undefined 并直接抛错
    //（官方文档写「两种写法等价」与实际实现不符，已实测确认）。
    comments: {
      children: doc.comments.map(
        (c): ICommentOptions => ({
          id: c.id,
          author: c.author || '管理员',
          date: new Date(c.date),
          ...(c.parentId !== undefined ? { parentId: c.parentId } : {}),
          ...(c.resolved !== undefined ? { resolved: c.resolved } : {}),
          children: [new Paragraph({ children: [new TextRun(c.text)] })],
        }),
      ),
    },
    sections,
  })
}

const DOCX_MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

/**
 * 打包成 docx，并补写「行」单位段距。
 *
 * 必须再过一遍 JSZip：docx 库不暴露 w:beforeLines / w:afterLines，只能在它
 * 序列化好的 styles.xml 上补（见 lineUnits.ts）。两个坑都实测过：
 *
 *   1. 改写内容时必须传 `createFolders: false` —— 默认会往包里塞一个目录条目
 *      （形如 `word/`），Word 打开带这种条目的 docx 会**卡死在打开动作上不返回**
 *      （2026-09-13，用 Word COM 复现）。纯解包再打包不受影响。
 *   2. 整包重新生成不改变「Word 能不能打开」这件事，纯解包再打包已实测通过。
 */
async function buildZip(doc: DocModel, spec: Spec, meta: ExportMeta): Promise<JSZip> {
  const raw = await Packer.toArrayBuffer(buildDocument(doc, spec, meta))
  const zip = await JSZip.loadAsync(raw)
  const styles = zip.file('word/styles.xml')
  if (styles) {
    const xml = await styles.async('string')
    zip.file('word/styles.xml', patchStylesXml(xml, lineUnitPlan(spec)), {
      createFolders: false,
    })
  }
  return zip
}

/** 打包成 Blob，浏览器里直接下载用 */
export async function toBlob(
  doc: DocModel,
  spec: Spec,
  meta: ExportMeta = {},
): Promise<Blob> {
  const zip = await buildZip(doc, spec, meta)
  return zip.generateAsync({ type: 'blob', mimeType: DOCX_MIME })
}

/** 打包成 base64，方便在 node 端写文件或做校验 */
export async function toBase64(
  doc: DocModel,
  spec: Spec,
  meta: ExportMeta = {},
): Promise<string> {
  const zip = await buildZip(doc, spec, meta)
  return zip.generateAsync({ type: 'base64' })
}
