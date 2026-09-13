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
  CommentRangeEnd,
  CommentRangeStart,
  CommentReference,
  DeletedTextRun,
  Document,
  DocumentGridType,
  Footer,
  InsertedTextRun,
  LineRuleType,
  PageBreak,
  PageNumber,
  Packer,
  Paragraph,
  TextRun,
  UnderlineType,
} from 'docx'
import type {
  ICommentOptions,
  IParagraphStyleOptions,
  ISectionOptions,
  ParagraphChild,
} from 'docx'
import JSZip from 'jszip'

import { STYLE_KEYS, lineSpacePt, ptToHalfPoints, ptToTwips } from '../spec'
import type { Align, LineRule, Spec, TextStyleSpec } from '../spec'
import type { Block, DocModel, PageBreakBlock, TextBlock } from '../types'
import { computeNumbering } from '../numbering'
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
      return LineRuleType.AUTO
  }
}

/**
 * 段前/段后按「行」换算成 twips。
 *
 * 基准是**文档网格行高**而不是本段行距（见 spec.ts 的 lineSpacePt）。这里的
 * before/after 是后备值：Word 认 beforeLines/afterLines（由 lineUnits.ts 补写），
 * 不支持这对属性的渲染器才退回到这里的磅值 —— 两者必须同源，否则「谁在用什么」
 * 会变成一个看不出来的分叉。
 *
 * 行距一律显式写出（含 auto → 单倍 240 twips）：正文的默认行距被定义成了
 * 固定值，样式若不写就会继承它 —— 页脚那类要单倍行距的样式会被撑高。
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
    line: s.lineRule === 'auto' ? 240 : ptToTwips(s.linePt),
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

/** 页码段落。字体字号对齐全部来自「页脚」样式（Word 内置样式名）。 */
function pageNumberParagraph(spec: Spec): Paragraph {
  return new Paragraph({
    style: spec.styles.footer.id,
    children: [new TextRun({ children: [PageNumber.CURRENT] })],
  })
}

function textBlockParagraph(
  block: TextBlock,
  spec: Spec,
  numbering: Map<string, string>,
  pageBreakBefore = false,
): Paragraph {
  const children: ParagraphChild[] = []

  const prefix = numbering.get(block.id)
  if (prefix) {
    // 编号按需求写成正文文字，与预览完全一致
    children.push(new TextRun({ text: prefix }))
  }

  for (const inline of block.inlines) {
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

  // 正文不挂样式：它就是 Word 的 Normal（内置「正文」），
  // 格式定义在 buildDocument 的 styles.default.document 上。
  const styleId = block.kind === 'body' ? undefined : spec.styles[block.kind].id
  return new Paragraph({
    ...(styleId ? { style: styleId } : {}),
    ...(pageBreakBefore ? { pageBreakBefore: true } : {}),
    children,
  })
}

interface SectionGroup {
  restartNumbering: boolean
  /** 本节里的块，含分页符（分节符本身不进组，它只负责切组） */
  blocks: (TextBlock | PageBreakBlock)[]
}

/** 按分节符把内容切成若干节。分节符在 Word 里意味着新起一页 + 独立的页码序列。 */
function groupSections(doc: DocModel): SectionGroup[] {
  const groups: SectionGroup[] = [{ restartNumbering: true, blocks: [] }]
  for (const block of doc.blocks) {
    if (block.t === 'sectionBreak') {
      groups.push({ restartNumbering: block.restartNumbering, blocks: [] })
      continue
    }
    groups[groups.length - 1]?.blocks.push(block)
  }
  return groups
}

/**
 * 把一节里的块变成段落。
 *
 * 分页符优先写成「段前分页」挂在它后面那一段上（`w:pageBreakBefore`）——
 * 这样不会像插一个空段落那样在页顶多留一个空行。
 *
 * 后面没有段落可挂时（紧跟分节符，或者干脆在节末/文末）必须退回到独立段落里
 * 写一个 `w:br w:type="page"`：**换页这件事不能丢**。否则「分页符 + 分节符」
 * 连在一起时，只剩分节符的换页生效，看上去就是「只分了一次页」。
 */
function sectionParagraphs(
  group: SectionGroup,
  spec: Spec,
  numbering: Map<string, string>,
): Paragraph[] {
  const children: Paragraph[] = []
  let breakBefore = false
  for (const block of group.blocks) {
    if (block.t === 'pageBreak') {
      breakBefore = true
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
  const sections: ISectionOptions[] = groupSections(doc).map((group) => {
    const children: Paragraph[] = sectionParagraphs(group, spec, numbering)
    if (children.length === 0) children.push(new Paragraph({}))

    return {
      properties: {
        page: {
          size: { width: spec.page.size.width, height: spec.page.size.height },
          margin: {
            top: spec.page.margin.top,
            right: spec.page.margin.right,
            bottom: spec.page.margin.bottom,
            left: spec.page.margin.left,
            header: spec.page.header,
            footer: spec.page.footer,
          },
          ...(group.restartNumbering ? { pageNumbers: { start: 1 } } : {}),
        },
        // 文档网格。Word 的「行」单位段距（w:beforeLines）以它的 linePitch 为基准，
        // 没有它 Word 会按一套我们控制不了的行高去算，段间距就对不上了。
        // 网格类型取 lines（对齐行网格）：行距写成固定值的段落不受网格影响，
        // 量出来的行盒仍与预览一致 —— 它在这里的作用只是给「1 行」定一个磅值。
        grid: {
          type: DocumentGridType.LINES,
          linePitch: ptToTwips(spec.page.gridLinePt),
        },
      },
      footers: { default: new Footer({ children: [pageNumberParagraph(spec)] }) },
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
