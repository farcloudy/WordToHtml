/**
 * 对外入口。
 *
 * 这个包做三件事，对应需求的三条主线：
 *   1. 类 md 字符串 ⇄ 文档模型（md/parse.ts、md/serialize.ts）
 *   2. 文档模型 → docx（docx/export.ts）
 *   3. 文档模型 → 分页预览 DOM（预览侧组件用）
 *
 * 组件本体在 src/components/WordPaper.vue，演示页在 src/App.vue。
 */

export {
  BLOCK_KINDS,
  DEFAULT_SPEC,
  contentBoxPx,
  lengthToPx,
  mm,
  ptToHalfPoints,
  ptToPx,
  ptToTwips,
  resolveSpec,
} from './spec'
export type {
  Align,
  BlockKind,
  DeepPartial,
  Length,
  LineRule,
  NumberingStyle,
  PageSpec,
  Spec,
  TextStyleSpec,
} from './spec'

export { collectRevisions, emptyDoc, nextBlockId, plainText, sliceInlines } from './types'
export type {
  Block,
  CommentDef,
  CommentEndInline,
  CommentStartInline,
  DocModel,
  Inline,
  RevMark,
  SectionBreakBlock,
  TextBlock,
  TextInline,
} from './types'

export { chineseNum, computeNumbering, numberingPrefix, stripAutoNumber } from './numbering'

export { parseMd } from './md/parse'
export type { ParseOptions } from './md/parse'
export { normalizeBlocks, toMd } from './md/serialize'

export { buildDocument, paragraphStyles, toBase64, toBlob } from './docx/export'
export type { ExportMeta } from './docx/export'

export { WTP, buildCss, injectCss } from './render/css'
export { escapeHtml, renderInlinesHtml } from './render/html'
export { measureDocument } from './render/measure'
export { paginate } from './render/paginate'
export type {
  MeasuredBlock,
  MeasuredBreak,
  MeasuredItem,
  PageFragment,
  PageLayout,
  PaginateOptions,
} from './render/paginate'
