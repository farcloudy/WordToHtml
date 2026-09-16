/**
 * 从规格表派生预览用的 CSS。
 *
 * 关键点：字体栈写成「西文字体在前、中文字体在后」—— 浏览器按字形逐个回退，
 * 拉丁字符命中 Times New Roman、中日韩字符回退到仿宋，正好复现 Word 里
 * w:rFonts 的 ascii / eastAsia 分工。这是预览能和 docx 对上的基础。
 */

import { STYLE_KEYS, lineSpacePt } from '../spec'
import type { Spec, TextStyleSpec } from '../spec'

export const WTP = 'wtp'

/**
 * 焦点离开正文时给选区「续命」用的高亮名（CSS Custom Highlight API）。
 * 名字在这里定义，注册与清除在编辑层做，两边必须用同一个。
 */
export const KEEP_SELECTION_HIGHLIGHT = `${WTP}-keep-selection`

/**
 * 查找命中与「当前命中」两层高亮的名字，同样走 CSS Custom Highlight API。
 * 匹配高亮绝不动 DOM —— 一个节点都不许改，否则就会违反「正常输入不得重排」。
 * 名字在这里定义，组件里注册与清除时引用同一份常量。
 */
export const SEARCH_HIGHLIGHT = `${WTP}-search-match`
export const SEARCH_CURRENT_HIGHLIGHT = `${WTP}-search-current`

/**
 * 整格复选（表格）的高亮类名。组件按模型坐标把它挂在命中的 `<td>` 上（见 render/html.ts 的
 * `data-cell-id`），这里只给一条底色规则 —— 这条规则**绝不能**动 padding / height / border：
 * 量测与预览共用 renderTableFragment，选中态要是改了盒模型，
 * 「量到的就是看到的」这个前提当场就不成立了。
 */
export const CELL_SELECTION_CLASS = `${WTP}-cellsel`

function quote(name: string): string {
  return `"${name.replace(/"/g, '')}"`
}

function fontStack(ascii: string, eastAsia: string): string {
  return `${quote(ascii)}, ${quote(eastAsia)}, serif`
}

function lineHeightOf(s: TextStyleSpec): string {
  return s.lineRule === 'auto' ? 'normal' : `${s.linePt}pt`
}

function textAlignOf(s: TextStyleSpec): string {
  return s.align === 'both' ? 'justify' : s.align
}

/** 注入到 document.head 的样式标签 id，同一个页面只保留一份 */
const STYLE_ID = 'wtp-preview-style'

/**
 * 生成整份预览 CSS。ns 允许改命名空间，避免和使用方的类名撞车。
 */
export function buildCss(spec: Spec, ns: string = WTP): string {
  const out: string[] = []

  out.push(
    /*
     * 页带：横排 + 可换行。容器够宽时两页并排，不够时自动回落到一页一排
     * ——「宽度足够就并排」这个条件由 flex-wrap 自己满足，不做缩放或适配页宽。
     * 打印不受影响：@media print 里把这里改回 display:block，仍是一页一张纸。
     */
    `.${ns}-pages { display: flex; flex-direction: row; flex-wrap: wrap; justify-content: center; align-items: flex-start; gap: 18px; }`,
    // flex: none（不放大也不缩小）：纸宽是绝对的，容器不够宽时宁可让外层横向滚动，
    // 也不能把纸压窄 —— 纸一窄版心就变，实测出来的行盒和分页算术立刻全体对不上。
    //
    // 纸宽/纸高/页边距（padding）**不在这里**：W5 起每节可以有各自的页面方向，
    // 它们是按页给的行内样式（见 WordPaper.vue 的 pageStyle）。这里只留与节无关的。
    `.${ns}-page { flex: none; position: relative; box-sizing: border-box; background: #fff; ` +
      `box-shadow: 0 1px 6px rgba(0, 0, 0, 0.18); }`,
    // 版心用 flex 纵向排列：flex 子项之间的 margin 不会合并，
    // 这一点很重要 —— Word 里段后与段前是相加的，而普通块级布局会取较大者。
    `.${ns}-content { display: flex; flex-direction: column; width: 100%; height: 100%; overflow: hidden; }`,
    // 页码元素只负责定位，字体字号对齐全部来自 .${ns}-footer（= 内置「页脚」样式），
    // 这样预览的页码排版与 docx 里那条样式同源。页脚距（bottom）也是逐节的，走行内样式。
    `.${ns}-page-number { position: absolute; left: 0; right: 0; color: #000; }`,
    // 测量容器：必须参与布局（不能用 display:none），否则量不到行盒
    `.${ns}-probe { position: absolute; left: -100000px; top: 0; visibility: hidden; pointer-events: none; }`,
    // 编辑态：版面本身就是编辑区，去掉浏览器默认的聚焦描边
    `.${ns}-content[contenteditable='true'] { outline: none; caret-color: #1f6feb; }`,
    // 自动编号是生成物，不该被选中或改到
    `.${ns}-num { -webkit-user-select: none; user-select: none; }`,
  )

  for (const kind of STYLE_KEYS) {
    const s = spec.styles[kind]
    out.push(
      `.${ns}-${kind} { font-family: ${fontStack(s.ascii, s.eastAsia)}; font-size: ${s.sizePt}pt; ` +
        `font-weight: ${s.bold ? 'bold' : 'normal'}; text-align: ${textAlignOf(s)}; ` +
        `line-height: ${lineHeightOf(s)}; text-indent: ${s.firstLineChars > 0 ? `${s.firstLineChars}em` : '0'}; ` +
        // 段前/段后按文档网格行高换算（lineSpacePt），与 docx 的 w:beforeLines 同源
        `margin: ${lineSpacePt(s.spaceBeforeLines, spec)}pt 0 ${lineSpacePt(s.spaceAfterLines, spec)}pt; ` +
        `white-space: pre-wrap; }`,
    )
  }

  /*
   * 表格。宽度铺满版心，固定布局、边框合并。
   *
   * 单元格左右内边距**必须**与 docx 侧一致：导出侧写死 108 缇 = 5.4pt（docx/export.ts 的
   * CELL_MARGIN_TWIPS），两侧不同，格内文字相对 Word 就会偏 —— 这不是巧合，是一个共用的数。
   *
   * 「最小一行 / 最小两行」的行高下限写在 **`<td>` 的 height 上**（表格格的 height 就是
   * 最小高度：内容更高时照样撑开，与 docx 侧 w:trHeight ATLEAST 同义）。**不能**写在格内
   * 那层 div 的 min-height 上 —— 那样 div 自己就被撑成两行高，而字仍待在 div 顶部，
   * 于是 `<td>` 上的 vertical-align 挪的是「已经填满格子的 div」，一点视觉效果都没有
   * （2026-09-15 用户实测：最小两行 + 单行文字时垂直对齐失效，设成居中/底端都不动）。
   *
   * 格内文字可以换成别的样式（`TableCellModel.kind`），所以这里按 STYLE_KEYS 逐条样式生成
   * —— 一行里各格取自己样式的那条，实际行高天然是各格的最大值
   * （docx 侧的 w:trHeight 用同一条规则算，两侧同源）。
   *
   * **表头行 / 附注行（`td.wtp-td-plain`）恒为一行高**（W7）：`minLines` 只管正文行，
   * 所以带 `-min2` 的那条规则把 plain 行排除在外，plain 行的下限由下面不带 `-min2` 的
   * 那条规则提供（1 行高）。两行的高度是「整张表的装饰」而不是内容，
   * 跟着行高设置一起变高只会白白把表格撑长（导出侧同一条规则，见 docx/export.ts 的 rowHeight）。
   *
   * atLeast 行距在预览里也写成固定值（CSS 的 line-height 只有固定值与 normal 两种），
   * 而 Word 的「最小值」是「不小于」：实际行高取 max(linePt, 字体自然行高)。两套模板的列表系列
   * 取值刻意取成「该字号的自然行高」（10.5pt → 12pt、14pt → 16pt；仿宋/黑体/楷体的自然行高
   * 倍数都是 1.1406 = 292/256），所以两边最多差 0.03pt。换字号或改这两个值时要把这层对上
   * （见 PLAN 第 5 节）。
   */
  out.push(
    `.${ns}-table { width: 100%; table-layout: fixed; border-collapse: collapse; }`,
    `.${ns}-table td { border: 0.5pt solid #000; padding: 0 5.4pt; vertical-align: top; ` +
      `overflow-wrap: break-word; }`,
    // unit / note 行整行一格且无边框（与导出侧的 NO_BORDERS 对应）
    `.${ns}-table td.${ns}-td-plain { border: 0; }`,
    // 整格复选的高亮：只给底色。类名由组件在 DOM 上按 data-cell-id 挂（见 CELL_SELECTION_CLASS），
    // 不进 renderTableFragment —— 那样会把选中态带进量测路径。
    `.${ns}-table td.${ns}-cellsel { background: rgba(31, 111, 235, 0.22); }`,
  )
  for (const kind of STYLE_KEYS) {
    out.push(
      `.${ns}-table td.${ns}-td-${kind} { height: ${spec.styles[kind].linePt}pt; }`,
      // :not(.wtp-td-plain) 是「表头行 / 附注行恒一行」的落点：那条规则在特异性上必须真的落空，
      // 否则 plain 行的 height 会被它盖掉（两条规则的类数不同，靠 :not 排除最稳）
      `.${ns}-table-min2 td.${ns}-td-${kind}:not(.${ns}-td-plain) { ` +
        `height: ${spec.styles[kind].linePt * 2}pt; }`,
    )
  }
  out.push(
    // 表格片段自己是普通块（内容宽 = 版心宽），刻意外层不挂 data-block-id（见 render/html.ts）
    `.${ns}-tableFrag { width: 100%; }`,
  )

  // 修订与批注的预览外观：是近似，不去追求和 Word 像素级一致
  out.push(
    `.${ns}-rev-ins { color: #1b7f3b; text-decoration: underline; text-decoration-color: #1b7f3b; }`,
    `.${ns}-rev-del { color: #b3261e; text-decoration: line-through; }`,
    `.${ns}-comment { background: rgba(255, 213, 0, 0.28); border-bottom: 1px dotted #a6791d; }`,
    // 侧栏里选中某条批注时，正文里的锚点加深，方便对上位置
    `.${ns}-comment-active { background: rgba(255, 213, 0, 0.62); }`,
    // 点批注输入框或工具栏时焦点离开正文，浏览器会把原生选区收起来 ——
    // 选区看起来就「丢了」。这里用 Custom Highlight 单独画一层：不动 DOM，
    // 因此不会碰坏正在编辑的内容，选区回到正文时再撤掉。
    `::highlight(${KEEP_SELECTION_HIGHLIGHT}) { background: rgba(31, 111, 235, 0.32); }`,
    // 查找命中：一层淡黄把所有命中铺出来，当前那一处再压一层更重的橙
    `::highlight(${SEARCH_HIGHLIGHT}) { background: rgba(255, 213, 0, 0.45); }`,
    `::highlight(${SEARCH_CURRENT_HIGHLIGHT}) { background: rgba(255, 145, 0, 0.75); }`,
  )

  /*
   * 打印：一张纸正好一页 .wtp-page，别的一概不出。
   *
   * @page 的尺寸取自规格表（纸张是可覆盖的，写死 A4 就错了）；margin 归 0 ——
   * 纸上的白边由 .wtp-page 自己的 padding（= 页边距）提供，页盒这边再留一次
   * 就成了双份，内容会被挤小。
   *
   * 分页不能靠 @page，我们的每页就是一个固定高度的 div：让每一页都在自己之前断页。
   * 用 break-before 而不是 break-after —— 页与页之间还夹着 .wtp-break（打印时
   * display:none，但仍是子元素），末尾那枚会让 :last-child 落空，「最后一页不再断」
   * 就失效并多印一张空白纸；:first-child 则不受影响（每页的第一个子元素就是它自己）。
   * 同时关掉屏幕上那些 flex / 阴影 / 间隙，flex 容器的子项断页行为不可靠。
   *
   * 全部带 !important：SFC 的 <style scoped> 会给选择器加上 data-v 属性，特异性
   * 高过这里的一层类名（.wtp-break[data-v-x] 会盖掉 .wtp-break），打印样式必须压得住它。
   */
  out.push(
    `@media print {`,
    // 未命名的那条是兜底（纸只有一档尺寸 = 规格表）；混排方向靠下面两条命名页规则，
    // 每页在行内挂 `page: wtp-portrait|wtp-landscape`（见 WordPaper.vue 的 pageStyle）。
    // landscape 那条把规格表的宽高互换 —— @page 的 size 是「纸的实际朝向」，不认 w:orient 那套。
    `  @page { size: ${spec.page.size.width} ${spec.page.size.height}; margin: 0; }`,
    `  @page ${ns}-portrait { size: ${spec.page.size.width} ${spec.page.size.height}; margin: 0; }`,
    `  @page ${ns}-landscape { size: ${spec.page.size.height} ${spec.page.size.width}; margin: 0; }`,
    `  html, body { background: #fff; }`,
    `  .${ns}-root { display: block !important; }`,
    `  .${ns}-pages { display: block !important; gap: 0 !important; }`,
    `  .${ns}-page { margin: 0 !important; box-shadow: none !important; ` +
      `overflow: hidden !important; break-before: page !important; }`,
    // 第一页就是第一张纸，前面不能再断一次（放在后面，压住上一条）
    `  .${ns}-page:first-child { break-before: auto !important; }`,
    // 批注侧栏、页间换页标记、隐藏的量测容器都不是纸上的东西
    `  .${ns}-comments, .${ns}-break, .${ns}-measure-root, .${ns}-probe { display: none !important; }`,
    `}`,
  )

  return out.join('\n')
}

/**
 * 把预览 CSS 注入 document.head（幂等）。
 * 之所以不用 SFC 的 <style>：样式内容依赖 props 传进来的规格，不是静态的。
 */
export function injectCss(spec: Spec, ns: string = WTP): void {
  if (typeof document === 'undefined') return
  let el = document.getElementById(STYLE_ID) as HTMLStyleElement | null
  if (!el) {
    el = document.createElement('style')
    el.id = STYLE_ID
    document.head.appendChild(el)
  }
  el.textContent = buildCss(spec, ns)
}
