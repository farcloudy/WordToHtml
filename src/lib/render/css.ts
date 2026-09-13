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
    `.${ns}-pages { display: flex; flex-direction: column; align-items: center; gap: 18px; }`,
    `.${ns}-page { position: relative; width: ${spec.page.size.width}; height: ${spec.page.size.height}; ` +
      `padding: ${spec.page.margin.top} ${spec.page.margin.right} ${spec.page.margin.bottom} ${spec.page.margin.left}; ` +
      `box-sizing: border-box; background: #fff; box-shadow: 0 1px 6px rgba(0, 0, 0, 0.18); }`,
    // 版心用 flex 纵向排列：flex 子项之间的 margin 不会合并，
    // 这一点很重要 —— Word 里段后与段前是相加的，而普通块级布局会取较大者。
    `.${ns}-content { display: flex; flex-direction: column; width: 100%; height: 100%; overflow: hidden; }`,
    // 页码元素只负责定位，字体字号对齐全部来自 .${ns}-footer（= 内置「页脚」样式），
    // 这样预览的页码排版与 docx 里那条样式同源。
    `.${ns}-page-number { position: absolute; left: 0; right: 0; bottom: ${spec.page.footer}; color: #000; }`,
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
    `  @page { size: ${spec.page.size.width} ${spec.page.size.height}; margin: 0; }`,
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
