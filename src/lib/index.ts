/**
 * 对外入口。
 *
 * 这个包做三件事，对应需求的三条主线：
 *   1. 类 md 字符串 ⇄ 文档模型（md/parse.ts、md/serialize.ts）
 *   2. 文档模型 → docx（docx/export.ts）
 *   3. 文档模型 → 分页预览 DOM（预览侧组件用）
 *
 * 组件：`WtpEditor` 是**对外的那个组件**（顶栏 + 功能区 + 纸张，见 components/WtpEditor.vue）；
 * `WordPaper` 是它内部的纸张组件（预览 + 编辑层），只做高级用法时才需要直接用。
 * 演示页在 src/App.vue（demo 专属的「类 md 源码」pane 与模式切换不会进这个入口）。
 */

export { default as WtpEditor } from '../components/WtpEditor.vue'
export { default as WordPaper } from '../components/WordPaper.vue'

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
  resolveEditorFlags,
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
  EditorFlags,
  EditorSettings,
  Inline,
  InlineHolder,
  PageBreakBlock,
  PageOrientation,
  RevMark,
  SectionBreakBlock,
  SectionSettings,
  TableBlock,
  TableCellAlign,
  TableCellModel,
  TableRowModel,
  TableRowRole,
  TextBlock,
  TextInline,
} from './types'

export { resolveSectionSettings, resolveSections } from './section'
export type { ResolvedSection, ResolvedSectionSettings, SectionRuntime } from './section'

export {
  insertSectionBreakAfter,
  normalizeSectionSettings,
  removeSectionBreak,
  sectionCountOf,
  sectionIndexOf,
  setSectionSetting,
  settingsOf,
} from './edit/section'

export { chineseNum, computeNumbering, numberingPrefix, stripAutoNumber } from './numbering'

export {
  addComment,
  applyFormat,
  blockLength,
  canJoinWithNext,
  canMergeIntoPrevious,
  cloneDoc,
  containerLength,
  deleteRange,
  deleteSpan,
  findBlock,
  findBlockIndex,
  findContainer,
  hasRevisions,
  insertBreakAfter,
  insertText,
  joinWithNext,
  mergeIntoPrevious,
  rangeColor,
  rangeIsBold,
  rangeIsUnderline,
  removeBlock,
  removeBreak,
  removeComment,
  replyComment,
  replaceRange,
  resolveRevisions,
  revisionSpanAt,
  setBlockKind,
  setContainerKind,
  sliceStrict,
  splitBlock,
  updateComment,
} from './edit/model'
export type {
  BlockPoint,
  CellSelectionSummary,
  EditorSelection,
  InlineContainer,
  RevisionSpan,
  SectionSelectionContext,
  TableSelectionContext,
} from './edit/model'

export {
  bodyInsertIndex,
  bodyRowIndexes,
  cellRectBetween,
  cellRectIndexOf,
  cellsChangingAlign,
  cellsChangingKind,
  cellsInRect,
  cellsInRects,
  findCell,
  findTable,
  insertBodyRow,
  insertColumn,
  nextAlignValue,
  normalizeCellCol,
  normalizeTable,
  removeBodyRow,
  removeColumn,
  removeTable,
  setCellAlign,
  setCellKind,
  setCellsAlign,
  setCellsKind,
  setMinLines,
  setRoleRow,
  sortCells,
  stepCell,
  storedCellAlign,
  storedCellKind,
  verticalCell,
} from './edit/table'
export type { CellRect, CellRef, CellStep } from './edit/table'

export { formatAmount } from './edit/amount'

export {
  DEFAULT_SHORTCUTS,
  SHORTCUT_ACTIONS,
  comboLabel,
  matchShortcut,
  parseCombo,
  resolveShortcuts,
} from './edit/shortcuts'
export type {
  Combo,
  ShortcutAction,
  ShortcutEvent,
  ShortcutOverrides,
  ShortcutTable,
} from './edit/shortcuts'

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
  CELL_SELECTION_CLASS,
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
