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
