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
  DEFAULT_MARGIN,
  DEFAULT_SPEC,
  DOC_TEMPLATES,
  MARGIN_PRESETS,
  STYLE_KEYS,
  contentBoxPx,
  lengthToPx,
  lineSpacePt,
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
  DocTemplate,
  Length,
  LineRule,
  MarginPreset,
  MarginSpec,
  NumberingStyle,
  PageSpec,
  Spec,
  StyleKey,
  TextStyleSpec,
} from './spec'

export {
  allInlineHolders,
  cellId,
  collectRevisions,
  commentScopes,
  defaultCellAlignH,
  emptyDoc,
  inlinesText,
  nextBlockId,
  parseCellId,
  plainText,
  sliceInlines,
} from './types'
export type {
  Block,
  BreakInline,
  CellVerticalAlign,
  CommentDef,
  CommentEndInline,
  CommentStartInline,
  DocModel,
  Inline,
  InlineHolder,
  PageBreakBlock,
  RevMark,
  SectionBreakBlock,
  TableBlock,
  TableCellAlign,
  TableCellModel,
  TableRowModel,
  TableRowRole,
  TextBlock,
  TextInline,
} from './types'

export { chineseNum, computeNumbering, numberingPrefix, stripAutoNumber } from './numbering'

export {
  addComment,
  applyFormat,
  blockLength,
  cloneDoc,
  containerLength,
  deleteRange,
  findBlock,
  findBlockIndex,
  findContainer,
  insertBreakAfter,
  insertText,
  mergeIntoPrevious,
  rangeColor,
  rangeIsBold,
  rangeIsUnderline,
  removeBlock,
  removeBreak,
  removeComment,
  replyComment,
  replaceRange,
  setBlockKind,
  setContainerKind,
  sliceStrict,
  splitBlock,
  updateComment,
} from './edit/model'
export type { BlockPoint, EditorSelection, InlineContainer, TableSelectionContext } from './edit/model'

export {
  bodyInsertIndex,
  bodyRowIndexes,
  findCell,
  findTable,
  insertBodyRow,
  insertColumn,
  normalizeTable,
  removeBodyRow,
  removeColumn,
  removeTable,
  setCellAlign,
  setCellKind,
  setMinLines,
  setRoleRow,
  stepCell,
  verticalCell,
} from './edit/table'
export type { CellStep } from './edit/table'

export { formatAmount } from './edit/amount'

export { findMatches, replaceMatches, validateQuery } from './edit/search'
export type { Match, SearchOptions, SearchScope } from './edit/search'

export { buildOutline, outlineSignature } from './edit/outline'
export type { OutlineEntry } from './edit/outline'

export {
  cssColorToHex,
  currentRange,
  displayPointOf,
  fragmentAt,
  fragmentOf,
  offsetToPoint,
  placeCaret,
  placeCaretAfterBreak,
  placeRange,
  pointToOffset,
  prefixLengthOf,
  readInlines,
  selectedRanges,
} from './edit/dom'
export type { DisplayPoint, DisplayRange } from './edit/dom'

export { parseMd } from './md/parse'
export type { ParseOptions } from './md/parse'
export { normalizeBlocks, toMd } from './md/serialize'

export { buildDocument, paragraphStyles, toBase64, toBlob } from './docx/export'
export type { ExportMeta } from './docx/export'
export { lineUnitPlan, patchStylesXml } from './docx/lineUnits'
export type { LineUnit, LineUnitPlan } from './docx/lineUnits'

export {
  KEEP_SELECTION_HIGHLIGHT,
  SEARCH_CURRENT_HIGHLIGHT,
  SEARCH_HIGHLIGHT,
  WTP,
  buildCss,
  injectCss,
} from './render/css'
export { escapeHtml, renderInlinesHtml, renderTableFragment } from './render/html'
export { clearMeasureCache, measureDocument } from './render/measure'
export type { MeasureCache, MeasureCacheEntry, MeasuredCacheValue } from './render/measure'
export { isTableFragment, paginate } from './render/paginate'
export type {
  BreakKind,
  MeasuredBlock,
  MeasuredBreak,
  MeasuredItem,
  MeasuredTableRow,
  PageBreakMark,
  PageFragment,
  PageLayout,
  PaginateOptions,
} from './render/paginate'
