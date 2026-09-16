/**
 * 类 md 源码 → 文档模型。
 *
 * 这是一套「特定模式」的简化语法，不是通用 Markdown：
 * 一个非空行就是一段，行与行之间不会合并成同一段（公文写作习惯是一行一段，
 * 用空行分隔反而要多敲一次回车）。空行只起分隔作用，本身不产出内容。
 *
 * 块级标记（行首）：
 *   #      文本标题
 *   ##     一级标题      → 自动编号 一、
 *   ###    二级标题      → 自动编号 （一）
 *   ####   三级标题      → 自动编号 1、
 *   @      抬头（正文但取消首行缩进）
 *   >>     落款（正文但右对齐）
 *   %      附件标记（基于正文：黑体、顶格、段前 0、段后 1 行）
 *   -      列表段落
 *   !      列表标题
 *   ---    分节符（单独一行，新起一页；后面可跟 kwargs 描述**它开启的那一节**）
 *   ===    分页符（单独一行，只强制换页，页码连续）
 *   :::table … :::  表格围栏块（见下）
 *
 * 节的设置（语义见 types.ts 的 SectionSettings；只写非默认值，双向可逆）：
 *   ::section …   文档首行描述**首节**：numbers=off / restart=on / orientation=landscape
 *   --- …         分节符所开启的节：link=off / numbers=off / restart=on / orientation=landscape
 *   例：`--- link=off restart=on`（独立页脚 + 页码从 1 重排，= 旧 `---` 的含义）
 *
 * 编辑器开关（语义见 types.ts 的 EditorSettings；同样只写非默认值、双向可逆）：
 *   ::editor …    文档开头的 `trackChanges=on` / `nav=off`（默认：修订关、导航开）
 *   与 `::section` 同在开头的指令区，谁先谁后都行；正文里出现的 `::editor` 落回正文。
 *
 * 表格围栏块：
 *   :::table minLines=2
 *   > 单位：元
 *   | 项目 | 金额 |
 *   | 甲 | 1,234.00 |
 *   < 注：以上金额不含税
 *   :::
 * 行首 `>` 是 unit 行、`<` 是 note 行（整行一格），`|` 是 body 行（按未转义的竖线切格）；
 * 其它行首忽略。kwarg 只写非默认值：minLines 默认 1、cantSplit 默认 true。
 * 格内可用 `{p}` 分段（Word 的单元格里放多个 `w:p`）：`| 甲{p}乙 |` 是一格两段。
 *
 * 行内标记：
 *   **文字**          加粗
 *   __文字__          下划线
 *   {红|文字}         改色，支持中文色名或 #RRGGBB
 *   {+文字}           插入修订
 *   {-文字}           删除修订
 *   [[文字|批注内容]]  批注，锚定在「文字」上
 *   {br}              软换行（Word 的 <w:br/>），零宽、不占字符位
 *   {p}               格内段落标记（**只在表格格里认**；正文里它不是语法）
 *   反斜杠 \ 转义上述所有标记字符
 *
 * 行内标记可以互相嵌套，例如 {红|**重点**}、{+**新增且加粗**}。
 */

import { BLOCK_KINDS } from '../spec'
import type { Align, BlockKind } from '../spec'
import type {
  Block,
  CellVerticalAlign,
  CommentDef,
  DocModel,
  EditorSettings,
  Inline,
  RevMark,
  SectionSettings,
  TableBlock,
  TableCellAlign,
  TableCellModel,
  TableRowModel,
} from '../types'
import { nextBlockId, resolveEditorFlags } from '../types'
import { emptyCell } from '../edit/table'
import { normalizeSectionSettings } from '../edit/section'

export interface ParseOptions {
  /** 修订与批注的作者名，写进 docx 的 w:author */
  author?: string
  /** 时间戳来源，测试里注入固定值以获得确定性输出 */
  now?: () => Date
  /** 修订 id 起始值 */
  startRevisionId?: number
}

const COLOR_NAMES: Record<string, string> = {
  红: 'FF0000',
  黑: '000000',
  蓝: '0000FF',
  绿: '008000',
  紫: '800080',
  橙: 'FF8C00',
  灰: '808080',
  黄: 'BF8F00',
  白: 'FFFFFF',
}

function normalizeColor(token: string): string | undefined {
  const t = token.trim()
  if (/^#?[0-9a-fA-F]{6}$/.test(t)) return t.replace(/^#/, '').toUpperCase()
  return COLOR_NAMES[t]
}

interface InlineContext {
  bold: boolean
  underline: boolean
  color?: string
  rev?: RevMark
  comments: CommentDef[]
  author: string
  now: () => Date
  nextRevisionId: () => number
  nextCommentId: () => number
}

/** 找到与 index 处字符配对的闭合标记，处理嵌套与转义 */
function findMatchingBrace(src: string, openIndex: number): number {
  let depth = 0
  for (let i = openIndex; i < src.length; i += 1) {
    const c = src[i]
    if (c === '\\') {
      i += 1
      continue
    }
    if (c === '{') depth += 1
    else if (c === '}') {
      depth -= 1
      if (depth === 0) return i
    }
  }
  return -1
}

function findClosing(src: string, from: number, token: string): number {
  let i = from
  while (i < src.length) {
    const hit = src.indexOf(token, i)
    if (hit < 0) return -1
    // 数一下 token 前面连续的反斜杠，奇数表示被转义
    let backslashes = 0
    let j = hit - 1
    while (j >= 0 && src[j] === '\\') {
      backslashes += 1
      j -= 1
    }
    if (backslashes % 2 === 0) return hit
    i = hit + token.length
  }
  return -1
}

/**
 * 找顶层（不在 {} 内、未被转义）的第一个竖线。
 * 指令与批注都用它分隔左右两半，这样 {#FF0000|文字} 里的竖线不会被误判成分隔符。
 */
function findTopLevelBar(src: string): number {
  let depth = 0
  let i = 0
  while (i < src.length) {
    const c = src[i]
    if (c === '\\') {
      i += 2
      continue
    }
    if (c === '{') depth += 1
    else if (c === '}') depth -= 1
    else if (c === '|' && depth === 0) return i
    i += 1
  }
  return -1
}

function makeRev(kind: 'ins' | 'del', ctx: InlineContext): RevMark {
  return {
    kind,
    id: ctx.nextRevisionId(),
    author: ctx.author,
    date: ctx.now().toISOString(),
  }
}

/** 解析一段花括号指令，返回它产出的 inline 序列 */
function parseDirective(inner: string, ctx: InlineContext): Inline[] {
  // {br} = 软换行（Word 的 <w:br/>）。
  //
  // 必须排在下面「没有竖线就当成字面量」那一支之前；也刻意不用 `\n` 那种写法 ——
  // parseInline 的第一条分支是「反斜杠吃掉下一个字符」，`\n` 会被解析成字面量 n，
  // 想绕开就得在那条通用转义分支之前另开一条特判。花括号指令没有这个坑：
  // 正文里真想写「{br}」时，escapeText 会把它序列化成 `\{br\}`，不会撞车。
  if (inner === 'br') return [{ t: 'break' }]

  const head = inner[0]
  if (head === '+') {
    return parseInline(inner.slice(1), { ...ctx, rev: makeRev('ins', ctx) })
  }
  if (head === '-') {
    return parseInline(inner.slice(1), { ...ctx, rev: makeRev('del', ctx) })
  }

  const bar = findTopLevelBar(inner)
  if (bar < 0) return literal(inner)
  const color = normalizeColor(inner.slice(0, bar))
  if (!color) return literal(inner)
  return parseInline(inner.slice(bar + 1), { ...ctx, color })
}

function literal(text: string): Inline[] {
  return text === '' ? [] : [{ t: 'text', text }]
}

/**
 * 反转 `serialize.ts` 的 escapeText：把 `\x` 还原成字面量 x。
 *
 * 批注内容是**原样存进模型、再原样写进 docx** 的（不像正文那样逐字符走 inline 解析），
 * 所以序列化时逃逸过的字符必须在解析时还原回来 —— 否则 `\|`、`\\` 这类内容每往返一次
 * 就多一层反斜杠，越滚越多。规则与 parseInline 处理转义一致：一次左到右扫描，反斜杠吃掉下一个字符。
 */
function unescapeText(src: string): string {
  let out = ''
  for (let i = 0; i < src.length; i += 1) {
    const c = src[i] ?? ''
    if (c === '\\' && i + 1 < src.length) {
      out += src[i + 1]
      i += 1
      continue
    }
    out += c
  }
  return out
}

function parseInline(src: string, ctx: InlineContext): Inline[] {
  const out: Inline[] = []
  let buffer = ''

  const flush = (): void => {
    if (buffer === '') return
    out.push({
      t: 'text',
      text: buffer,
      ...(ctx.bold ? { bold: true } : {}),
      ...(ctx.underline ? { underline: true } : {}),
      ...(ctx.color ? { color: ctx.color } : {}),
      ...(ctx.rev ? { rev: ctx.rev } : {}),
    })
    buffer = ''
  }

  let i = 0
  while (i < src.length) {
    const c = src[i]

    if (c === '\\' && i + 1 < src.length) {
      buffer += src[i + 1]
      i += 2
      continue
    }

    if (src.startsWith('**', i)) {
      const close = findClosing(src, i + 2, '**')
      if (close >= 0 && close > i + 2) {
        flush()
        out.push(...parseInline(src.slice(i + 2, close), { ...ctx, bold: true }))
        i = close + 2
        continue
      }
    }

    if (src.startsWith('__', i)) {
      const close = findClosing(src, i + 2, '__')
      if (close >= 0 && close > i + 2) {
        flush()
        out.push(...parseInline(src.slice(i + 2, close), { ...ctx, underline: true }))
        i = close + 2
        continue
      }
    }

    if (c === '{') {
      const close = findMatchingBrace(src, i)
      if (close > i + 1) {
        flush()
        out.push(...parseDirective(src.slice(i + 1, close), ctx))
        i = close + 1
        continue
      }
    }

    if (src.startsWith('[[', i)) {
      const close = src.indexOf(']]', i + 2)
      if (close > i + 2) {
        const inner = src.slice(i + 2, close)
        const bar = findTopLevelBar(inner)
        if (bar > 0) {
          flush()
          const id = ctx.nextCommentId()
          ctx.comments.push({
            id,
            author: ctx.author,
            date: ctx.now().toISOString(),
            text: unescapeText(inner.slice(bar + 1)),
          })
          out.push({ t: 'commentStart', commentId: id })
          out.push(...parseInline(inner.slice(0, bar), ctx))
          out.push({ t: 'commentEnd', commentId: id })
          i = close + 2
          continue
        }
      }
    }

    buffer += c
    i += 1
  }

  flush()
  return out
}

interface BlockRule {
  re: RegExp
  kind: BlockKind
}

// 顺序要紧：长标记必须排在它的前缀之前（#### 先于 ### 先于 ## 先于 #）
const BLOCK_RULES: readonly BlockRule[] = [
  { re: /^####\s+(.*)$/, kind: 'h3' },
  { re: /^###\s+(.*)$/, kind: 'h2' },
  { re: /^##\s+(.*)$/, kind: 'h1' },
  { re: /^#\s+(.*)$/, kind: 'title' },
  { re: /^@\s*(.*)$/, kind: 'salutation' },
  { re: /^>>\s*(.*)$/, kind: 'signature' },
  { re: /^%\s+(.*)$/, kind: 'attachment' },
  { re: /^!\s+(.*)$/, kind: 'listTitle' },
  { re: /^-\s+(.*)$/, kind: 'listItem' },
]

const SECTION_BREAK_RE = /^-{3,}(?:\s+(.*))?$/
const PAGE_BREAK_RE = /^={3,}\s*$/
/**
 * 文档首行描述**首节**的指令：`::section numbers=off orientation=landscape`。
 * 只用两个冒号，与表格围栏的三个冒号不撞车；也只认文件第一处非空行，
 * 免得正文里写「::section …」被吃掉。
 */
const SECTION_DIRECTIVE_RE = /^::\s*section\b(.*)$/

/**
 * 文档开头的编辑器开关：`::editor trackChanges=on nav=on`（只写非默认值）。
 *
 * 与 `::section` 同住「开头的指令区」：两者都在文档最前面认，谁先谁后都行
 * （序列化固定写 `::editor` 在前、`::section` 在后，但手工写的 md 不该因此被吃掉）。
 * 一旦出现正文，指令区就关闭 —— 正文里写「::editor …」落回正文，与「围栏没闭合绝不当表格」
 * 同一条原则。
 */
const EDITOR_DIRECTIVE_RE = /^::\s*editor\b(.*)$/

/**
 * 解析分节 kwargs：`link=on|off` / `numbers=on|off` / `restart=on|off` / `orientation=portrait|landscape`。
 * 认不出的键与值一律忽略（不报错）；认出默认值也不落字段 —— 归一化交给
 * normalizeSectionSettings（与 setSectionSetting 同一个约定，模型里不留冗余值）。
 */
function parseSectionKwargs(attrs: string, isFirst: boolean): SectionSettings {
  const out: SectionSettings = {}
  for (const m of attrs.matchAll(/([A-Za-z]+)\s*=\s*(\S+)/g)) {
    const key = m[1]
    const value = m[2]
    if (key === 'numbers') {
      if (value === 'off') out.pageNumbers = false
      else if (value === 'on') out.pageNumbers = true
    } else if (key === 'link') {
      if (value === 'off') out.linkPrevious = false
      else if (value === 'on') out.linkPrevious = true
    } else if (key === 'restart') {
      if (value === 'on') out.restartAtOne = true
      else if (value === 'off') out.restartAtOne = false
    } else if (key === 'orientation') {
      if (value === 'landscape' || value === 'portrait') out.orientation = value
    }
  }
  normalizeSectionSettings(out, isFirst)
  return out
}

/**
 * 解析 `::editor` 的 kwargs：`trackChanges=on|off` / `nav=on|off`。
 *
 * 只认这两个键与 on/off 两个值，其余一律忽略（不报错）；认出**默认值也不落字段** ——
 * 默认值不落字段是「模型 → md → 模型」字节稳定的前提（与分节 kwargs 同一套约定）。
 * 两个键全是默认值时返回 undefined（整个 editor 字段都不写进模型）。
 */
function parseEditorKwargs(attrs: string): EditorSettings | undefined {
  const out: EditorSettings = {}
  for (const m of attrs.matchAll(/([A-Za-z]+)\s*=\s*(\S+)/g)) {
    const key = m[1]
    const value = m[2]
    if (key === 'trackChanges') {
      if (value === 'on') out.trackChanges = true
      else if (value === 'off') out.trackChanges = false
    } else if (key === 'nav') {
      if (value === 'on') out.nav = true
      else if (value === 'off') out.nav = false
    }
  }
  if (out.trackChanges === undefined && out.nav === undefined) return undefined
  const flags = resolveEditorFlags(out)
  const normalized: EditorSettings = {}
  if (flags.trackChanges) normalized.trackChanges = true
  if (!flags.nav) normalized.nav = false
  return Object.keys(normalized).length > 0 ? normalized : undefined
}

/** 表格围栏的起始行：`:::table`，后面可跟 minLines=1|2、headerRows=N、cantSplit=yes|no */
const TABLE_FENCE_RE = /^:::\s*table\b(.*)$/
/** 表格围栏的结束行（trim 后逐字比较） */
const TABLE_FENCE_END = ':::'

/** 解析围栏行后面的 kwarg；只认 minLines / headerRows / cantSplit，其余忽略 */
function parseTableFence(attrs: string): { minLines: 1 | 2; cantSplit: boolean; headerRows: number } {
  let minLines: 1 | 2 = 1
  let cantSplit = true
  let headerRows = 0
  for (const m of attrs.matchAll(/([A-Za-z]+)\s*=\s*(\S+)/g)) {
    const key = m[1]
    const value = m[2]
    if (key === 'minLines' && value === '2') minLines = 2
    else if (key === 'minLines' && value === '1') minLines = 1
    else if (key === 'cantSplit' && value === 'yes') cantSplit = true
    else if (key === 'cantSplit' && value === 'no') cantSplit = false
    else if (key === 'headerRows') headerRows = Number.parseInt(value ?? '', 10)
  }
  return { minLines, cantSplit, headerRows }
}

/** src[i] 处是不是被反斜杠转义（前面连续的反斜杠个数为奇数） */
function isEscapedAt(src: string, i: number): boolean {
  let n = 0
  for (let j = i - 1; j >= 0 && src[j] === '\\'; j -= 1) n += 1
  return n % 2 === 1
}

/** 剥掉紧邻的一个空格（单元格两侧的空白就是靠这一步与解析时的一个空格配对） */
function stripOneSpace(s: string): string {
  return s.startsWith(' ') ? s.slice(1) : s
}

/**
 * 按**顶层**的 `|` 把一行 body 切成若干格。
 *
 * 「顶层」= 没有落在 `{}` 指令、`[[]]` 批注、或反斜杠转义里面 —— 格内允许写
 * `{红|甲}`、`[[甲|核对原件]]` 这类行内语法，它们内部的竖线是内容的一部分，
 * 不算分隔符（只看反斜杠是不够的，那正是最初漏掉 `{…}` / `[[]]` 的成因）。
 *
 * 取法是「先剥掉行首那枚外框 `|`，再剥掉每格紧邻的一个空格，最后丢掉行尾那枚
 * 外框留下的空尾格」—— 与序列化写出的 `| a | b |` 严格互逆，所以 ` a`、`a `
 * 这种带空格的内容也能往返。
 */
function splitTableCells(line: string): string[] {
  let s = line
  if (s.startsWith('|')) s = s.slice(1)

  const raw: string[] = []
  let cur = ''
  let depth = 0
  let comment = false

  for (let i = 0; i < s.length; i += 1) {
    const ch = s[i] ?? ''
    if (comment) {
      cur += ch
      if (ch === ']' && s[i + 1] === ']' && !isEscapedAt(s, i)) {
        cur += ']'
        i += 1
        comment = false
      }
      continue
    }
    if (ch === '[' && s[i + 1] === '[' && !isEscapedAt(s, i)) {
      cur += '[['
      i += 1
      comment = true
      continue
    }
    if (ch === '{' && !isEscapedAt(s, i)) depth += 1
    else if (ch === '}' && depth > 0 && !isEscapedAt(s, i)) depth -= 1
    if (ch === '|' && depth === 0 && !isEscapedAt(s, i)) {
      raw.push(cur)
      cur = ''
      continue
    }
    cur += ch
  }
  raw.push(cur)

  // 行尾那枚外框 `|` 会留下一个空尾格，丢掉它（整行只有一枚 `|` 时保留成单空格）
  if (raw.length > 1 && raw[raw.length - 1] === '') raw.pop()
  if (raw.length === 0) raw.push('')

  return raw.map((cell) => {
    let t = stripOneSpace(cell)
    if (t.endsWith(' ')) t = t.slice(0, -1)
    return t
  })
}

const CELL_H_ALIGNS: readonly string[] = ['left', 'center', 'right']
const CELL_V_ALIGNS: readonly string[] = ['top', 'middle', 'bottom']

/** 格内段落标记 */
const CELL_PARA_MARK = '{p}'

/**
 * 按**顶层**的 `{p}` 把一格的内容切成若干段。
 *
 * 「顶层」= 没落在 `{}` 指令、`[[]]` 批注、反斜杠转义里面 —— 与 splitTableCells 同一套
 * depth / comment 扫描。`{p}` 只在表格格里认（**不并进通用的 parseDirective**，否则正文里
 * 也会长出这套语义）：`| 甲{p}乙 |` 是一格两段，`{br}` 仍是段内的软换行。
 *
 * 首段为空（`| {p}乙 |`）与尾随空段（`| 甲{p} |`）都**保留成真实的一段**：它们是模型里
 * 合法的 `[[], [文字]]` / `[[文字], []]`，往返必须双向可逆、字节稳定。
 */
function splitCellParagraphs(text: string): string[] {
  const out: string[] = []
  let cur = ''
  let depth = 0
  let comment = false

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i] ?? ''
    if (comment) {
      cur += ch
      if (ch === ']' && text[i + 1] === ']' && !isEscapedAt(text, i)) {
        cur += ']'
        i += 1
        comment = false
      }
      continue
    }
    if (ch === '[' && text[i + 1] === '[' && !isEscapedAt(text, i)) {
      cur += '[['
      i += 1
      comment = true
      continue
    }
    if (ch === '{' && !isEscapedAt(text, i)) {
      // 只有「顶层且正好是完整的 {p}」才算段落标记；`\{p\}` 里的那个被反斜杠吃掉，不算
      if (depth === 0 && text.startsWith(CELL_PARA_MARK, i)) {
        out.push(cur)
        cur = ''
        i += CELL_PARA_MARK.length - 1
        continue
      }
      depth += 1
    } else if (ch === '}' && depth > 0 && !isEscapedAt(text, i)) {
      depth -= 1
    }
    cur += ch
  }
  out.push(cur)
  return out
}

/** 一格的内容切成 kind / align 前缀与正文（正文里可能还带 `{p}`，由调用方再切段） */
function splitCellAttrPrefix(text: string): { kind?: BlockKind; align?: TableCellAlign; body: string } {
  if (!text.startsWith('{@')) return { body: text }

  const close = findMatchingBrace(text, 0)
  if (close < 0) return { body: text }
  const inner = text.slice(2, close)
  const bar = findTopLevelBar(inner)
  if (bar < 0) return { body: text }

  let kind: BlockKind | undefined
  let align: TableCellAlign | undefined
  for (const raw of inner.slice(0, bar).split(',')) {
    const token = raw.trim()
    if (BLOCK_KINDS.includes(token as BlockKind)) {
      kind = token as BlockKind
    } else if (CELL_H_ALIGNS.includes(token)) {
      align = { ...(align ?? {}), h: token as Align }
    } else if (CELL_V_ALIGNS.includes(token)) {
      align = { ...(align ?? {}), v: token as CellVerticalAlign }
    }
  }

  // 闭合 `}` 之后若还有内容，一并当正文（正常写法不会有，但不许静默丢掉）
  return {
    ...(kind !== undefined ? { kind } : {}),
    ...(align !== undefined ? { align } : {}),
    body: inner.slice(bar + 1) + text.slice(close + 1),
  }
}

/**
 * 单元格内容 → 模型。开头可选指令 `{@<token>[,<token>]…|<格内正文>}`（本波之前就有的语法），
 * 格内正文再按顶层 `{p}` 切成**多段**（每段一段 inline）。
 *
 * 三类 token 互不冲突：BlockKind 名 → 该格样式；left/center/right → 水平对齐；
 * top/middle/bottom → 垂直对齐。不认识的 token 一律忽略、不报错；
 * 显式写出默认样式 `listItem` 同样不落字段（见函数末尾的注释）。
 *
 * 转义可逆：正文里真写 `{@x|y}` 会被 serialize.ts 的 escapeText 转成 `\{@x|y\}`，
 * 解析时反斜杠吃掉 `{`，就还原成字面量。
 * `{@` 之后找不到闭合的 `}`（或没有分隔 token 与正文的 `|`）时**当普通内容处理**，
 * 绝不吞掉后面的内容。
 */
function parseCell(text: string, ctx: InlineContext): TableCellModel {
  const { kind, align, body } = splitCellAttrPrefix(text)
  const cell: TableCellModel = {
    paragraphs: splitCellParagraphs(body).map((part) => ({ inlines: parseInline(part, ctx) })),
  }
  // 'listItem' 就是默认样式：认出它也不落字段。setCellKind 在设回默认时会删字段、序列化也会省略它，
  // 这里跟这两处对齐（模型里不留冗余值）—— 围栏的 minLines=1 / cantSplit 早就是这个约定。
  if (kind !== undefined && kind !== 'listItem') cell.kind = kind
  if (align !== undefined) cell.align = align
  return cell
}

/**
 * 解析一个 `:::table` 围栏块（含结束行），返回模型块与结束行的下标。
 *
 * **找不到结束行时返回 null**：半截围栏不成立 —— 既不能把后文整篇当表格吃掉
 *（那是静默丢内容），也不该假装表格已经写完。调用方会退回「按普通行继续解」。
 */
function parseTableBlock(
  lines: readonly string[],
  startIndex: number,
  attrs: string,
  ctx: InlineContext,
): { block: TableBlock; end: number } | null {
  const { minLines, cantSplit, headerRows } = parseTableFence(attrs)
  const rows: TableRowModel[] = []
  let closed = false

  let i = startIndex + 1
  for (; i < lines.length; i += 1) {
    const line = lines[i] ?? ''
    if (line.trim() === TABLE_FENCE_END) {
      closed = true
      break
    }
    if (line.trim() === '') continue

    if (line.startsWith('>') || line.startsWith('<')) {
      // 整行一格：文字就是标记之后的内容（再剥一个空格，与序列化写的 "> xxx" 互逆）。
      // 这一支**不能 trimEnd** —— 行尾空格属于内容；body 行不同，它的内容后面还有
      // 一枚外框 `|`，所以对 body 行 trimEnd 只会削掉框外的空白。
      const cells: TableCellModel[] = [parseCell(stripOneSpace(line.slice(1)), ctx)]
      rows.push({ role: line.startsWith('>') ? 'unit' : 'note', cells })
      continue
    }
    if (line.startsWith('|')) {
      rows.push({
        role: 'body',
        cells: splitTableCells(line.trimEnd()).map((text) => parseCell(text, ctx)),
      })
      continue
    }
    // 其它行首一律忽略：不报错、不产出内容
  }

  if (!closed) return null
  if (rows.length === 0) rows.push({ role: 'body', cells: [emptyCell()] })
  let columns = 1
  for (const row of rows) {
    if (row.role === 'body') columns = Math.max(columns, row.cells.length)
  }

  // headerRows 夹到 [0, 行数]：写 0 / 负数 / 超过行数一律收口，归零时不落字段
  // （「默认值不落模型」的既有约定）。夹到行数就是「整张表都是标题行」这个退化态，
  // 它在 Word 里是「每一页都把整张表重复一遍」，渲染 / 分页 / 导出按同一个 N 走。
  const header = Math.max(0, Math.min(Number.isFinite(headerRows) ? headerRows : 0, rows.length))

  return {
    block: {
      t: 'table',
      id: nextBlockId('tb'),
      rows,
      columns,
      minLines,
      cantSplit,
      ...(header > 0 ? { headerRows: header } : {}),
    },
    end: i,
  }
}


export function parseMd(source: string, options: ParseOptions = {}): DocModel {
  const author = options.author ?? '管理员'
  const now = options.now ?? (() => new Date())
  let revisionSeq = options.startRevisionId ?? 0
  let commentSeq = 0

  const comments: CommentDef[] = []
  const blocks: Block[] = []
  // 逐节设置，下标 = 节号。首节由文档首行的 ::section 指令描述；每遇到一个 `---`
  // 就往后面追加一项（它开启的那一节）。全默认时最后不写进模型（见函数末尾）。
  const sections: SectionSettings[] = [{}]
  let editor: EditorSettings | undefined
  /**
   * 还在「文档开头的指令区」里吗？`::editor` 与 `::section` 都只在这段里认，
   * 出现第一行正文就关闭 —— 免得正文里写「::editor …」被吃掉。
   */
  let inHeader = true

  const baseCtx = {
    bold: false,
    underline: false,
    comments,
    author,
    now,
    nextRevisionId: () => {
      const id = revisionSeq
      revisionSeq += 1
      return id
    },
    nextCommentId: () => {
      const id = commentSeq
      commentSeq += 1
      return id
    },
  }

  const lines = source.split(/\r?\n/)
  for (let li = 0; li < lines.length; li += 1) {
    const line = (lines[li] ?? '').trimEnd()
    if (line.trim() === '') continue

    // 文档开头的指令区：`::editor …`（编辑器开关）与 `::section …`（首节设置）都在这里认，
    // 谁先谁后都行；第一行正文一到就关闭（下面的 inHeader = false）。
    if (inHeader) {
      const editorHead = EDITOR_DIRECTIVE_RE.exec(line)
      if (editorHead) {
        editor = parseEditorKwargs(editorHead[1] ?? '')
        continue
      }
      const head = SECTION_DIRECTIVE_RE.exec(line)
      if (head) {
        sections[0] = parseSectionKwargs(head[1] ?? '', true)
        continue
      }
      inHeader = false
    }

    const fence = TABLE_FENCE_RE.exec(line)
    if (fence) {
      const parsed = parseTableBlock(lines, li, fence[1] ?? '', baseCtx)
      // 围栏没闭合（parsed 为 null）：不当表格 —— 围栏行按普通行继续解，绝不吞掉后文
      if (parsed) {
        blocks.push(parsed.block)
        li = parsed.end
        continue
      }
    }

    const sectionBreak = SECTION_BREAK_RE.exec(line)
    if (sectionBreak) {
      // `---` 后面可以跟 kwargs（只写非默认值），描述**它开启的**那一节
      sections.push(parseSectionKwargs(sectionBreak[1] ?? '', false))
      blocks.push({ t: 'sectionBreak', id: nextBlockId('s') })
      continue
    }

    if (PAGE_BREAK_RE.test(line)) {
      blocks.push({ t: 'pageBreak', id: nextBlockId('pg') })
      continue
    }

    let matched = false
    for (const rule of BLOCK_RULES) {
      const m = rule.re.exec(line)
      const body = m?.[1]
      if (body === undefined) continue
      blocks.push({
        t: 'textBlock',
        id: nextBlockId(),
        kind: rule.kind,
        inlines: parseInline(body, baseCtx),
      })
      matched = true
      break
    }
    if (matched) continue

    blocks.push({
      t: 'textBlock',
      id: nextBlockId(),
      kind: 'body',
      inlines: parseInline(line, baseCtx),
    })
  }

  // 全默认就不写这个字段：`sections: [{}, {}]` 与「没有 sections」语义相同，
  // 两种形态都能存在的话，「模型 → md → 模型」会在两者之间漂移（见 edit/section.ts）。
  const doc: DocModel = { blocks, comments }
  if (sections.some((s) => Object.keys(s).length > 0)) doc.sections = sections
  // editor 同理：全默认时 parseEditorKwargs 已经返回 undefined，这里只落非默认值
  if (editor) doc.editor = editor
  return doc
}
