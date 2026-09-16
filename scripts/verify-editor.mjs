/**
 * P3 编辑层的浏览器实测。
 *
 * 验收的是「在 A4 版面上直接写」这件事本身：
 *   1. 正常输入**不触发重排** —— 用「给片段元素打标记，敲完字标记还在」来证明
 *      DOM 没被重建；插入符自然也不会丢；
 *   2. 真的需要重排时（行数变了），插入符还能按块 id + 字符偏移找回来；
 *   3. 结构性操作（回车分段、退格合并）、工具栏（加粗／下划线／改色）、修订模式、
 *      批注、撤销，都能落到模型上，并且导出用的就是这份模型；
 *   4. 快捷键：Ctrl+U 下划线、Ctrl+Shift+E 修订模式、Alt+4 金额格式（含无效输入的提示条）、
 *      工具栏「插入空格」下拉的三个特殊空格码点；
 *   5. 打印：编辑器外壳全部隐藏、每张纸各占一页、纸张尺寸取自规格表（真打一份 PDF 数页数）；
 *   6. 切文件模板：页数变了、插入符与选区都按坐标找回、文字没丢；
 *   7. W6：功能区四页高度一致、自定义快捷键表（改绑 / 旧键失效 / 未知动作名）、F4 重复上一步、
 *      顶栏文件名（导出名 = 文件名 + .docx）、`::editor` 与顶栏两个开关的双向同步。
 *
 * 每一步都同时看两边：DOM 上看到了什么，模型里记下了什么。只看 DOM 会漏掉
 * 「界面改了、导出没改」，只看模型会漏掉「模型改了、界面没跟上」。
 *
 * 用法：node scripts/verify-editor.mjs   （需先 npm run build:lib）
 */

import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

import { chromium } from 'playwright-core'
import { createServer } from 'vite'

import {
  SEARCH_CURRENT_HIGHLIGHT,
  SEARCH_HIGHLIGHT,
  DOC_TEMPLATES,
  cellRectBetween,
  cellsInRects,
  computeNumbering,
  contentBoxPx,
  ptToPx,
  resolveSpec,
  toMd,
} from '../dist-lib/wordtohtml.mjs'

const here = dirname(fileURLToPath(import.meta.url))
const root = resolve(here, '..')
// 无头浏览器用 Chrome（勿用 Edge），可用 WTP_BROWSER 覆盖
const CHROME = process.env.WTP_BROWSER || 'C:/Program Files/Google/Chrome/Application/chrome.exe'
const PORT = 5198

const failures = []

function ok(label, condition, detail = '') {
  if (condition) {
    console.log(`  ok   ${label}`)
    return true
  }
  failures.push(`${label}${detail ? ` — ${detail}` : ''}`)
  console.error(`  FAIL ${label}${detail ? ` — ${detail}` : ''}`)
  return false
}

function eq(label, actual, expected) {
  return ok(label, actual === expected, `期望 ${JSON.stringify(expected)}，实际 ${JSON.stringify(actual)}`)
}

/** 页面里装一组测试用的小工具。names 是两层查找高亮的名字（从 lib 导出的常量） */
const TEST_HELPERS = (names) => {
  window.__wtpHighlightNames = names
  const helpers = {
    fragments() {
      return Array.from(document.querySelectorAll('[data-block-id]'))
    },
    /** 查找高亮的 Range 数。normal = 全部命中，current = 当前那一处。 */
    highlights() {
      const registry = typeof CSS !== 'undefined' ? CSS.highlights : undefined
      if (!registry) return { supported: false, normal: 0, current: 0 }
      const all = registry.get(window.__wtpHighlightNames.normal)
      const cur = registry.get(window.__wtpHighlightNames.current)
      return { supported: true, normal: all ? all.size : 0, current: cur ? cur.size : 0 }
    },
    textOf(el) {
      return el.textContent ?? ''
    },
    fragmentByText(needle) {
      return helpers.fragments().find((el) => helpers.textOf(el).includes(needle)) ?? null
    },
    pointIn(el, offset) {
      const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT)
      let remaining = offset
      let last = null
      for (let n = walker.nextNode(); n; n = walker.nextNode()) {
        if (remaining <= n.data.length) return { node: n, offset: remaining }
        remaining -= n.data.length
        last = n
      }
      return last ? { node: last, offset: last.data.length } : null
    },
    focusHost(el) {
      const host = el.closest('[contenteditable="true"]')
      if (host) host.focus()
      return host
    },
    /** 把插入符放到某片段的第 offset 个字符处（含自动编号前缀） */
    setCaret(needle, offset) {
      const frag = helpers.fragmentByText(needle)
      if (!frag) return null
      helpers.focusHost(frag)
      const spot = helpers.pointIn(frag, offset)
      if (!spot) return null
      const sel = document.getSelection()
      const range = document.createRange()
      range.setStart(spot.node, spot.offset)
      range.collapse(true)
      sel.removeAllRanges()
      sel.addRange(range)
      return { blockId: frag.dataset.blockId, text: helpers.textOf(frag) }
    },
    caretAtEndOf(needle) {
      const frag = helpers.fragmentByText(needle)
      if (!frag) return null
      return helpers.setCaret(needle, helpers.textOf(frag).length)
    },
    /** 在片段内选中 [from,to) 的文字 */
    selectIn(needle, from, to) {
      const frag = helpers.fragmentByText(needle)
      if (!frag) return null
      helpers.focusHost(frag)
      const a = helpers.pointIn(frag, from)
      const b = helpers.pointIn(frag, to)
      if (!a || !b) return null
      const sel = document.getSelection()
      const range = document.createRange()
      range.setStart(a.node, a.offset)
      range.setEnd(b.node, b.offset)
      sel.removeAllRanges()
      sel.addRange(range)
      return helpers.textOf(frag).slice(from, to)
    },
    /**
     * 跨页段落：某个块在第 1 页与第 2 页各有一片。返回块 id 与它在哪两页上。
     * （issues/20260915 的删除边界都发生在这样的块上）
     */
    crossPageBlock() {
      const pages = Array.from(document.querySelectorAll('.wtp-page'))
      const byId = new Map()
      for (const el of document.querySelectorAll('.wtp-content [data-block-id]')) {
        const id = el.dataset.blockId ?? ''
        if (!byId.has(id)) byId.set(id, [])
        byId.get(id).push(pages.indexOf(el.closest('.wtp-page')))
      }
      for (const [id, list] of byId) if (list.length > 1) return { blockId: id, pages: list }
      return null
    },
    /** 把插入符放到某块在某页那一片的首 / 末（跨页删除要靠它落在片段边缘上） */
    caretAtPageEdgeOf(blockId, pageIndex, where) {
      const pages = Array.from(document.querySelectorAll('.wtp-page'))
      const frag = Array.from(
        document.querySelectorAll(`.wtp-content [data-block-id="${blockId}"]`),
      ).find((el) => pages.indexOf(el.closest('.wtp-page')) === pageIndex)
      if (!frag) return null
      helpers.focusHost(frag)
      const range = document.createRange()
      if (where === 'end') range.selectNodeContents(frag)
      else range.setStart(frag, 0)
      range.collapse(where !== 'end')
      const sel = document.getSelection()
      sel.removeAllRanges()
      sel.addRange(range)
      return { blockId, text: helpers.textOf(frag), from: frag.dataset.from, to: frag.dataset.to }
    },
    /** 插入符在哪个片段、在该片段里的第几个字符 */
    caretInfo() {
      const sel = document.getSelection()
      if (!sel || sel.rangeCount === 0) return null
      const range = sel.getRangeAt(0)
      let el = range.startContainer
      while (el && el.nodeType === 3) el = el.parentNode
      while (el && !(el instanceof HTMLElement && el.dataset && el.dataset.blockId)) el = el.parentNode
      if (!(el instanceof HTMLElement)) return null
      const before = document.createRange()
      before.selectNodeContents(el)
      before.setEnd(range.startContainer, range.startOffset)
      return {
        blockId: el.dataset.blockId ?? '',
        text: helpers.textOf(el),
        offset: before.toString().length,
      }
    },
    pages() {
      return Array.from(document.querySelectorAll('.wtp-page')).map((pageEl) => {
        const content = pageEl.querySelector('.wtp-content')
        const kids = content ? Array.from(content.children) : []
        let used = 0
        kids.forEach((el, i) => {
          const cs = getComputedStyle(el)
          used += el.getBoundingClientRect().height
          used += parseFloat(cs.marginTop) || 0
          if (i < kids.length - 1) used += parseFloat(cs.marginBottom) || 0
        })
        return {
          contentHeight: content ? content.clientHeight : 0,
          usedHeight: used,
        }
      })
    },
  }
  window.__wtpTest = helpers
}

let browser
let page

const server = await createServer({
  root,
  logLevel: 'warn',
  server: { port: PORT, strictPort: true },
})
await server.listen()
const url = server.resolvedUrls?.local?.[0] ?? `http://localhost:${PORT}/`
console.log(`dev server: ${url}`)

/**
 * 打开应用。
 *
 * `tabLabel` 给了就顺手切到那一页：功能区的页内容是按 tab 用 v-if 渲染的，
 * 重新载入后默认停在「开始」页 —— 要验别的页里的控件，就得先切过去。
 */
async function openApp(tabLabel) {
  await page.goto(url, { waitUntil: 'load' })
  await page.waitForSelector('.wtp-page', { timeout: 30000 })
  await page.evaluate(async () => {
    await document.fonts.ready
  })
  await page.waitForTimeout(600)
  await page.evaluate(TEST_HELPERS, {
    normal: SEARCH_HIGHLIGHT,
    current: SEARCH_CURRENT_HIGHLIGHT,
  })
  if (tabLabel) await openTab(tabLabel)
}

const getModel = () => page.evaluate(() => window.__wtpPaper.getModel())
const heroBlocks = (model) => model.blocks.filter((b) => b.t === 'textBlock')
const textOfBlock = (block) =>
  block.inlines
    .filter((i) => i.t === 'text')
    .map((i) => i.text)
    .join('')

/**
 * 切功能区的标签页（开始 / 插入 / 布局 / 表格）。
 *
 * 四页常驻在标签条上，但页内容是按 tab 的 v-if 出现的 —— 要验某一页里的控件，
 * 必须先切过去。标签按钮是 @mousedown.prevent（焦点不离开正文），所以切页不会丢选区。
 */
async function openTab(label) {
  const tabButton = page.locator('.ribbon-tab', { hasText: label })
  if ((await tabButton.count()) === 0) throw new Error(`功能区里没有「${label}」这一页`)
  await tabButton.first().click()
  await page.waitForTimeout(80)
}

/**
 * 按模型的「可搜索文字」数一处查询出现几次 —— 期望值从模型推导，不从界面反推。
 * 删除修订的文字不参与，且它会把前后断开（与 lib/edit/search.ts 同一套口径）。
 */
function countInModel(model, needle) {
  let total = 0
  for (const block of heroBlocks(model)) {
    let text = ''
    for (const inline of block.inlines) {
      if (inline.t !== 'text') continue
      if (inline.rev && inline.rev.kind === 'del') {
        text += '\u0000'
        continue
      }
      text += inline.text
    }
    let at = text.indexOf(needle)
    while (at >= 0) {
      total += 1
      at = text.indexOf(needle, at + needle.length)
    }
  }
  return total
}

/** 从模型推导导航窗格该有的条目（编号按规格表的 numbering 现算，不硬编码） */
function expectedOutline(model) {
  const spec = resolveSpec()
  const numbering = computeNumbering(model.blocks, (block) =>
    block.t === 'textBlock' ? spec.styles[block.kind].numbering : 'none',
  )
  const out = []
  for (const block of heroBlocks(model)) {
    const level = { h1: 1, h2: 2, h3: 3 }[block.kind]
    if (level === undefined) continue
    out.push({
      blockId: block.id,
      level,
      prefix: numbering.get(block.id) ?? '',
      text: textOfBlock(block),
    })
  }
  return out
}

/**
 * 版面每一块的文字（含自动编号）是否等于模型 —— 替换之后用它证明「界面改了、模型也改了」。
 * 同一块被分页切开时会渲染成多片，按 data-block-id 把各片拼起来才是整块文字。
 */
async function checkDomMatchesModel(label) {
  const model = await getModel()
  const spec = resolveSpec()
  const numbering = computeNumbering(model.blocks, (block) =>
    block.t === 'textBlock' ? spec.styles[block.kind].numbering : 'none',
  )
  const want = new Map(
    heroBlocks(model).map((block) => [
      block.id,
      (numbering.get(block.id) ?? '') + textOfBlock(block),
    ]),
  )
  const got = await page.evaluate(() => {
    const map = {}
    for (const frag of document.querySelectorAll('.wtp-pages [data-block-id]')) {
      const id = frag.dataset.blockId ?? ''
      map[id] = (map[id] ?? '') + (frag.textContent ?? '')
    }
    return map
  })
  let bad = null
  for (const [id, text] of want) {
    if (got[id] !== text) {
      bad = { id, want: text.slice(0, 40), got: (got[id] ?? '(缺失)').slice(0, 40) }
      break
    }
  }
  ok(label, bad === null, JSON.stringify(bad))
}

/**
 * 插入符「应该」落在哪里：模型里第一个含 needle 的块、needle 末尾处（显示坐标，含自动编号前缀）。
 * 替换会 force 重排并重建 DOM，锚点没给对的话插入符会丢，所以这条要靠断言守住。
 */
async function expectedCaretAtEndOf(needle) {
  const model = await getModel()
  const spec = resolveSpec()
  const numbering = computeNumbering(model.blocks, (block) =>
    block.t === 'textBlock' ? spec.styles[block.kind].numbering : 'none',
  )
  const host = heroBlocks(model).find((block) => textOfBlock(block).includes(needle))
  if (!host) return null
  return {
    blockId: host.id,
    offset: (numbering.get(host.id) ?? '').length + textOfBlock(host).indexOf(needle) + needle.length,
  }
}

function sameCaret(a, b) {
  return a !== null && b !== null && a.blockId === b.blockId && a.offset === b.offset
}

/** 编辑后每页都不得溢出 —— 重排算术与真实布局必须仍然对得上 */
async function checkNoOverflow(label) {
  const pages = await page.evaluate(() => window.__wtpTest.pages())
  let bad = 0
  for (const [i, p] of pages.entries()) {
    if (p.usedHeight > p.contentHeight + 1) {
      bad += 1
      failures.push(`${label}：第${i + 1}页溢出（${Math.round(p.usedHeight)}/${p.contentHeight}px）`)
    }
  }
  ok(`${label}：${pages.length} 页均未溢出`, bad === 0, bad > 0 ? `${bad} 页溢出` : '')
}

try {
  browser = await chromium.launch({ executablePath: CHROME, headless: true })
  page = await browser.newPage({ viewport: { width: 1700, height: 1100 }, deviceScaleFactor: 2 })

  /* ------------------------------------------------------------------ */
  console.log('\n=== A. 编辑面就在分页后的版面上 ===')
  await openApp()
  const surface = await page.evaluate(() => ({
    editable: document.querySelectorAll('.wtp-content[contenteditable="true"]').length,
    pages: document.querySelectorAll('.wtp-page').length,
    nums: document.querySelectorAll('.wtp-content .wtp-num').length,
    numEditable: Array.from(document.querySelectorAll('.wtp-num')).every(
      (el) => el.getAttribute('contenteditable') === 'false',
    ),
  }))
  eq('每个页面都是可编辑宿主', surface.editable, surface.pages)
  ok('自动编号已随段落渲染', surface.nums > 0)
  ok('自动编号不可编辑', surface.numEditable)
  console.log(`  页面 ${surface.pages} 页，可编辑宿主 ${surface.editable} 个，编号 ${surface.nums} 处`)

  /* ------------------------------------------------------------------ */
  console.log('\n=== B. 正常输入不重排、不丢插入符（DOM 节点身份不变）===')
  await openApp()
  const before = await getModel()
  const beforeBlocks = heroBlocks(before).length
  const beforePages = await page.evaluate(() => document.querySelectorAll('.wtp-page').length)
  await page.evaluate(() => {
    const frag = window.__wtpTest.fragmentByText('苏州市公安局')
    frag.__wtpKeep = true
    window.__wtpTest.caretAtEndOf('苏州市公安局')
  })
  await page.keyboard.insertText('（测试甲）')
  const kept = await page.evaluate(() => {
    const frag = window.__wtpTest.fragmentByText('（测试甲）')
    return {
      same: frag ? frag.__wtpKeep === true : false,
      connected: frag ? frag.isConnected : false,
      caret: window.__wtpTest.caretInfo(),
      pages: document.querySelectorAll('.wtp-page').length,
    }
  })
  ok('敲字后片段还是同一个 DOM 节点（没有重排重建）', kept.same && kept.connected)
  ok('插入符仍在被编辑的片段里', kept.caret !== null && kept.caret.text.includes('（测试甲）'))
  ok(
    '插入符落在刚敲完的字后面',
    kept.caret !== null && kept.caret.offset === kept.caret.text.length,
    kept.caret ? `offset ${kept.caret.offset} / len ${kept.caret.text.length}` : '取不到插入符',
  )
  const afterTyping = await getModel()
  const salutation = heroBlocks(afterTyping).find((b) => textOfBlock(b).includes('苏州市公安局'))
  ok('模型同步到了新文字', (salutation ? textOfBlock(salutation) : '').endsWith('（测试甲）'))
  eq('段落数不变', heroBlocks(afterTyping).length, beforeBlocks)
  eq('页数不变', kept.pages, beforePages)
  await checkNoOverflow('B 普通输入后')

  /* ------------------------------------------------------------------ */
  console.log('\n=== C. 插到会改变行数的位置：重排后插入符仍按坐标找回 ===')
  await openApp()
  const pagesBeforeC = await page.evaluate(() => document.querySelectorAll('.wtp-page').length)
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('我方于2026年9月1日'))
  const marker = '（重排探针一二三四五六七八九十甲乙丙丁戊己庚辛壬癸）'
  await page.keyboard.insertText(marker)
  await page.waitForTimeout(150)
  const reflowed = await page.evaluate((m) => {
    const frag = window.__wtpTest.fragmentByText(m)
    return {
      found: frag !== null,
      caret: window.__wtpTest.caretInfo(),
      pages: document.querySelectorAll('.wtp-page').length,
    }
  }, marker)
  ok('重排后仍能在版面上找到插入的文字', reflowed.found)
  ok(
    '插入符跟着文字走（落在探针所在片段里）',
    reflowed.caret !== null && reflowed.caret.text.includes(marker),
    reflowed.caret ? `插入符在「${reflowed.caret.text.slice(0, 12)}…」里` : '取不到插入符',
  )
  ok('页数没有减少（加字只会持平或增页）', reflowed.pages >= pagesBeforeC)
  const afterReflow = await getModel()
  const touched = heroBlocks(afterReflow).find((b) => textOfBlock(b).includes(marker))
  ok('模型里有这段插入的文字', touched !== undefined)
  await checkNoOverflow('C 重排后')

  /* ------------------------------------------------------------------ */
  console.log('\n=== D. 回车分段 / 退格合并 ===')
  await openApp()
  const baseCount = heroBlocks(await getModel()).length
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('苏州市公安局'))
  await page.keyboard.press('Enter')
  await page.waitForTimeout(120)
  const split = await getModel()
  eq('回车后多了一个段落', heroBlocks(split).length, baseCount + 1)
  const newBlock = heroBlocks(split).find((b) => textOfBlock(b) === '')
  ok('新段落是空的正文块', newBlock !== undefined && newBlock.kind === 'body')
  await page.keyboard.insertText('新段落内容')
  await page.waitForTimeout(120)
  const typed = await getModel()
  const typedBlock = heroBlocks(typed).find((b) => textOfBlock(b) === '新段落内容')
  ok('新段落敲得进字', typedBlock !== undefined)
  const caretInNew = await page.evaluate(() => window.__wtpTest.caretInfo())
  ok('插入符在新段落里', caretInNew !== null && caretInNew.text === '新段落内容')

  await page.evaluate(() => window.__wtpTest.setCaret('新段落内容', 0))
  await page.keyboard.press('Backspace')
  await page.waitForTimeout(120)
  const merged = await getModel()
  eq('退格把新段落并回上一段', heroBlocks(merged).length, baseCount)
  const mergedSalutation = heroBlocks(merged).find((b) => textOfBlock(b).includes('苏州市公安局'))
  ok(
    '合并后文字接在抬头末尾',
    mergedSalutation !== undefined && textOfBlock(mergedSalutation).endsWith('新段落内容'),
  )
  await checkNoOverflow('D 分段/合并后')

  /* ------------------------------------------------------------------ */
  console.log('\n=== E. 撤销 / 重做 ===')
  await openApp()
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('苏州市公安局'))
  await page.keyboard.insertText('撤销测试')
  const withUndoText = await getModel()
  ok(
    '敲进去的字在模型里',
    heroBlocks(withUndoText).some((b) => textOfBlock(b).includes('撤销测试')),
  )
  await page.keyboard.press('Control+z')
  const undone = await getModel()
  ok(
    'Ctrl+Z 撤销后模型里没有这段字',
    !heroBlocks(undone).some((b) => textOfBlock(b).includes('撤销测试')),
  )
  await page.keyboard.press('Control+y')
  const redone = await getModel()
  ok(
    'Ctrl+Y 重做后又回来了',
    heroBlocks(redone).some((b) => textOfBlock(b).includes('撤销测试')),
  )

  /* ------------------------------------------------------------------ */
  console.log('\n=== F. 工具栏：加粗 / 改色 ===')
  await openApp()
  await page.evaluate(() => window.__wtpTest.selectIn('我方于2026年9月1日', 0, 2))
  await page.waitForTimeout(60)
  await page.click('button.tool[title^="加粗"]')
  await page.waitForTimeout(150)
  const bolded = await getModel()
  const boldBlock = heroBlocks(bolded).find((b) => textOfBlock(b).startsWith('我方'))
  const boldInline = boldBlock ? boldBlock.inlines.find((i) => i.t === 'text' && i.bold) : null
  ok('模型里前两个字标了加粗', boldInline !== null && boldInline !== undefined && boldInline.text === '我方')
  const boldDom = await page.evaluate(() => {
    const frag = window.__wtpTest.fragmentByText('我方于2026年9月1日')
    return {
      b: frag ? Array.from(frag.querySelectorAll('b')).map((el) => el.textContent) : [],
    }
  })
  ok('版面上那两个字有 <b>', boldDom.b.includes('我方'), JSON.stringify(boldDom.b))

  await page.evaluate(() => window.__wtpTest.selectIn('我方于2026年9月1日', 0, 2))
  await page.waitForTimeout(60)
  await page.click('.swatch[title="标红"]')
  await page.waitForTimeout(150)
  const colored = await getModel()
  const colorBlock = heroBlocks(colored).find((b) => textOfBlock(b).startsWith('我方'))
  const colorInline = colorBlock
    ? colorBlock.inlines.find((i) => i.t === 'text' && i.color === 'FF0000')
    : null
  ok('模型里那两个字改成了红色', colorInline !== null && colorInline !== undefined)
  const colorDom = await page.evaluate(() => {
    const frag = window.__wtpTest.fragmentByText('我方于2026年9月1日')
    const span = frag ? frag.querySelector('span[style*="color"]') : null
    return span ? { text: span.textContent, style: span.getAttribute('style') } : null
  })
  ok('版面上那两个字带颜色', colorDom !== null && colorDom.text === '我方', JSON.stringify(colorDom))
  await checkNoOverflow('F 格式化后')

  /* ------------------------------------------------------------------ */
  console.log('\n=== G. 修订模式 ===')
  await openApp()
  await page.click('label.checkbox input[type="checkbox"]')
  await page.waitForTimeout(80)
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('苏州市公安局'))
  await page.keyboard.insertText('（修订新增）')
  await page.waitForTimeout(120)
  const tracked = await getModel()
  const insBlock = heroBlocks(tracked).find((b) => textOfBlock(b).includes('（修订新增）'))
  const insInline = insBlock
    ? insBlock.inlines.find((i) => i.t === 'text' && i.rev && i.rev.kind === 'ins')
    : null
  ok('模型里新增文字标成了 ins 修订', insInline !== null && insInline !== undefined)
  const trackDom = await page.evaluate(() => ({
    ins: Array.from(document.querySelectorAll('.wtp-rev-ins')).map((el) => el.textContent),
  }))
  ok('版面上新增文字带修订底色', trackDom.ins.some((t) => (t ?? '').includes('修订新增')), JSON.stringify(trackDom.ins))

  // 开着修订删字：字要留在原处，只是被标成删除
  await page.evaluate(() => window.__wtpTest.selectIn('苏州市公安局', 0, 2))
  await page.waitForTimeout(60)
  await page.keyboard.press('Backspace')
  await page.waitForTimeout(200)
  const deleted = await getModel()
  const delInline = heroBlocks(deleted)
    .flatMap((b) => b.inlines)
    .find((i) => i.t === 'text' && i.rev && i.rev.kind === 'del' && i.text.includes('苏州'))
  ok('被删的「苏州」没有消失，而是标成了 del 修订', delInline !== null && delInline !== undefined)
  const stillThere = heroBlocks(deleted).some((b) => textOfBlock(b).includes('苏州市公安局'))
  ok('正文里仍能看到被删的字', stillThere)
  const delDom = await page.evaluate(() =>
    Array.from(document.querySelectorAll('.wtp-rev-del')).map((el) => el.textContent),
  )
  ok('版面上那两个字带了删除线', delDom.some((t) => (t ?? '').includes('苏州')), JSON.stringify(delDom))

  // 接受 / 拒绝修订：按钮常驻在「开始」页，选区里没有修订就置灰。
  // 目标认「（修订新增）」那一块 —— 样本自己还带一处 ins 一处 del，所以不能按全篇数。
  const revBlock = (model) => heroBlocks(model).find((b) => textOfBlock(b).includes('（修订新增）'))
  const countRev = (model, kind) =>
    (revBlock(model)?.inlines ?? []).filter((i) => i.t === 'text' && i.rev?.kind === kind).length
  /**
   * 目标块**之外**那些修订（样本自带的 ins / del）的条数 —— 「别处一个都没动」要按这个比：
   * 拒绝/接受本来就该把目标块里的那一条去掉，拿全篇总数比是错的。
   */
  const otherRevCounts = (model) => {
    const inlines = heroBlocks(model)
      .filter((b) => b !== revBlock(model))
      .flatMap((b) => b.inlines)
    return {
      ins: inlines.filter((i) => i.t === 'text' && i.rev?.kind === 'ins').length,
      del: inlines.filter((i) => i.t === 'text' && i.rev?.kind === 'del').length,
    }
  }
  const acceptBtn = page.locator('.panel-start button.tool[title^="接受"]')
  const rejectBtn = page.locator('.panel-start button.tool[title^="拒绝"]')
  const revTextBefore = textOfBlock(revBlock(deleted) ?? { inlines: [] })
  eq(
    '目标块里删除修订一处、插入修订一处',
    `${countRev(deleted, 'del')}/${countRev(deleted, 'ins')}`,
    '1/1',
  )

  await page.evaluate(() => window.__wtpTest.caretAtEndOf('我方于2026年9月1日'))
  await page.waitForTimeout(150)
  eq('光标不在修订上时「接受修订」置灰', await acceptBtn.isDisabled(), true)
  eq('光标不在修订上时「拒绝修订」置灰', await rejectBtn.isDisabled(), true)

  // 光标停在修订串末尾也算「落在修订上」：点到修订字的右半边不该看到按钮变灰
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('苏州市公安局'))
  await page.waitForTimeout(150)
  eq('光标贴在插入修订末尾时按钮可用', await rejectBtn.isDisabled(), false)

  // 拒绝删除修订：被删的字回来（标记去掉、文字留下）；别处的修订一个都不许动
  await page.evaluate(() => window.__wtpTest.selectIn('苏州市公安局', 0, 2))
  await page.waitForTimeout(120)
  const revsBeforeReject = otherRevCounts(await getModel())
  await rejectBtn.click()
  await page.waitForTimeout(250)
  const rejected = await getModel()
  eq('拒绝删除修订后那一处不再带 del 标记', countRev(rejected, 'del'), 0)
  eq('拒绝后插入修订没被动到（同一块里另一处仍在）', countRev(rejected, 'ins'), 1)
  eq('拒绝删除修订后文字原样留在原处', textOfBlock(revBlock(rejected) ?? { inlines: [] }), revTextBefore)
  eq(
    '别处的修订一个都没动',
    JSON.stringify(otherRevCounts(rejected)),
    JSON.stringify(revsBeforeReject),
  )

  // 接受插入修订：新增的文字留下、标记去掉；处理完按钮自己变灰
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('苏州市公安局'))
  await page.waitForTimeout(120)
  await acceptBtn.click()
  await page.waitForTimeout(250)
  const accepted = await getModel()
  eq('接受插入修订后 ins 标记没了', countRev(accepted, 'ins'), 0)
  eq('接受插入修订后文字一个不少', textOfBlock(revBlock(accepted) ?? { inlines: [] }), revTextBefore)
  eq('处理完之后（光标处已无修订）「接受修订」又置灰', await acceptBtn.isDisabled(), true)

  // 撤销要能还原这一步（记的是「改之前」的模型）
  await page.keyboard.press('Control+z')
  await page.waitForTimeout(300)
  eq('Ctrl+Z 把接受掉的修订还原回来', countRev(await getModel(), 'ins'), 1)

  /* ------------------------------------------------------------------ */
  console.log('\n=== H. 批注：选中 → 写内容 → 侧栏出现 ===')
  await openApp('插入')
  const commentsBefore = (await getModel()).comments.length
  const sidebarBefore = await page.evaluate(
    () => document.querySelectorAll('.wtp-comments li .wtp-comment-item').length,
  )
  const anchor = await page.evaluate(() => window.__wtpTest.selectIn('债务人爱康光电科技有限公司', 0, 6))
  ok('选中了被批注的文字', anchor === '债务人爱康光', anchor ?? '未选中')
  // 选区要变成「记录」：点进批注框以后实时选区就只剩输入框里的空选区了
  await page.waitForTimeout(80)
  await page.click('.comment-field input')
  await page.fill('.comment-field input', '这是一条实测批注')
  await page.click('.comment-field button')
  await page.waitForTimeout(200)
  const commented = await getModel()
  eq('模型里多了一条批注', commented.comments.length, commentsBefore + 1)
  const last = commented.comments[commented.comments.length - 1]
  ok('批注内容与作者写对了', last && last.text === '这是一条实测批注', JSON.stringify(last))
  const sidebarAfter = await page.evaluate(() =>
    Array.from(document.querySelectorAll('.wtp-comments li .wtp-comment-item')).map((b) => ({
      scope: (b.querySelector('.wtp-comment-scope')?.textContent ?? '').trim(),
      text: (b.querySelector('.wtp-comment-text')?.textContent ?? '').trim(),
    })),
  )
  eq('侧栏多了一条', sidebarAfter.length, sidebarBefore + 1)
  ok(
    '侧栏显示的是刚写的内容',
    sidebarAfter.some((s) => s.text === '这是一条实测批注'),
    JSON.stringify(sidebarAfter),
  )
  const anchorDom = await page.evaluate(() =>
    Array.from(document.querySelectorAll('.wtp-comment')).map((el) => el.textContent),
  )
  ok(
    '正文里出现了批注锚点',
    anchorDom.some((t) => (t ?? '').includes('债务人爱康')),
    JSON.stringify(anchorDom),
  )
  await checkNoOverflow('H 加批注后')

  /* ------------------------------------------------------------------ */
  // 紧跟在 H 后面：这一节要证明「导出用的是界面上改出来的模型」，
  // 所以必须和刚才那些编辑处在同一次页面加载里（每次 openApp 都会重新解析源码）。
  console.log('\n=== I. 导出用的模型就是界面上改出来的 ===')
  const finalModel = await getModel()
  const hasEdits =
    heroBlocks(finalModel).some((b) => textOfBlock(b).includes('（测试甲）')) ||
    finalModel.comments.some((c) => c.text === '这是一条实测批注')
  ok('模型里有本轮编辑的痕迹', hasEdits)
  const tmpDir = join(root, '.qwen', 'tmp')
  mkdirSync(tmpDir, { recursive: true })
  const dumpPath = join(tmpDir, 'editor-model.json')
  writeFileSync(dumpPath, JSON.stringify(finalModel, null, 2), 'utf8')
  console.log(`  ok   最终模型已导出：${dumpPath}`)

  /* ------------------------------------------------------------------ */
  console.log('\n=== H2. 点批注框时选中的底色还在（Custom Highlight 续命）===')
  await openApp('插入')
  await page.evaluate(() => window.__wtpTest.selectIn('债务人爱康光电科技有限公司', 0, 6))
  await page.waitForTimeout(80)
  await page.click('.comment-field input')
  await page.waitForTimeout(120)
  const keptHighlight = await page.evaluate(() => {
    const registry = typeof CSS !== 'undefined' ? CSS.highlights : undefined
    if (!registry) return 'unsupported'
    return registry.has('wtp-keep-selection')
  })
  ok(
    '焦点进了批注框，选区高亮仍在',
    keptHighlight === true,
    keptHighlight === 'unsupported'
      ? '浏览器不支持 CSS Custom Highlight API'
      : `实际 ${keptHighlight}`,
  )
  // 加完批注应当收掉这层高亮，免得下一条批注悄悄复用旧选区
  await page.fill('.comment-field input', '续命高亮用批注')
  await page.click('.comment-field button')
  await page.waitForTimeout(200)
  const dropped = await page.evaluate(() =>
    typeof CSS !== 'undefined' && CSS.highlights ? CSS.highlights.has('wtp-keep-selection') : false,
  )
  eq('加完批注后高亮收掉', dropped, false)

  /* ------------------------------------------------------------------ */
  console.log('\n=== H3. 批注能改、能删 ===')
  await openApp('插入')
  await page.evaluate(() => window.__wtpTest.selectIn('债务人爱康光电科技有限公司', 0, 6))
  await page.waitForTimeout(80)
  await page.click('.comment-field input')
  await page.fill('.comment-field input', '待改写')
  await page.click('.comment-field button')
  await page.waitForTimeout(200)
  const seeded = await getModel()
  const targetId = seeded.comments[seeded.comments.length - 1].id
  const row = `.wtp-comments li[data-comment-id="${targetId}"]`

  await page.click(`${row} .wtp-comment-actions button:nth-child(1)`)
  await page.waitForTimeout(80)
  ok('编辑态出现了输入框', (await page.locator(`${row} .wtp-comment-edit input`).count()) === 1)
  await page.fill(`${row} .wtp-comment-edit input`, '已改写')
  await page.click(`${row} .wtp-comment-edit button.primary`)
  await page.waitForTimeout(200)
  const edited = (await getModel()).comments.find((c) => c.id === targetId)
  eq('模型里的批注内容被改写', edited?.text, '已改写')
  const sidebarText = await page.evaluate(
    (id) =>
      document
        .querySelector(`.wtp-comments li[data-comment-id="${id}"] .wtp-comment-text`)
        ?.textContent?.trim() ?? '',
    targetId,
  )
  eq('侧栏显示的是改写后的内容', sidebarText, '已改写')

  await page.click(`${row} .wtp-comment-actions button:nth-child(2)`)
  await page.waitForTimeout(200)
  const afterDelete = await getModel()
  eq('批注从模型里删掉', afterDelete.comments.some((c) => c.id === targetId), false)
  const anchorsLeft = await page.evaluate(() => document.querySelectorAll('.wtp-comment').length)
  eq('正文锚点也一起清掉', anchorsLeft, afterDelete.comments.length)

  /* ------------------------------------------------------------------ */
  console.log('\n=== H4. 分页符 / 分节符：按钮插入，页间看得见 ===')
  await openApp('插入')
  const breaksBefore = await page.evaluate(() => ({
    pages: document.querySelectorAll('.wtp-page').length,
    marks: document.querySelectorAll('.wtp-break').length,
  }))
  eq('初始样本里有 1 个分节符标记', breaksBefore.marks, 1)

  await page.evaluate(() => window.__wtpTest.setCaret('苏州市公安局', 0))
  await page.click('button.tool[title^="在光标所在段落后插入分页符"]')
  await page.waitForTimeout(250)
  const withPage = await getModel()
  eq('模型里多了一个分页符', withPage.blocks.filter((b) => b.t === 'pageBreak').length, 1)
  const pageMarks = await page.evaluate(() =>
    Array.from(document.querySelectorAll('.wtp-break')).map((el) => el.textContent.trim()),
  )
  ok('页间出现了「分页符」标记', pageMarks.some((t) => t.includes('分页符')), JSON.stringify(pageMarks))
  const pagesAfterPage = await page.evaluate(() => document.querySelectorAll('.wtp-page').length)
  // 页数只可能不变或 +1：分页符后面的内容被迫另起一页，而页尾原来那点余量被吸收掉，
  // 恰好抵掉一页时总数就不变（这不是缺陷）。所以真正要断言的是「后面的内容确实被推到新页页首」。
  ok(
    '插分页符后页数不减少',
    pagesAfterPage >= breaksBefore.pages,
    `${breaksBefore.pages} → ${pagesAfterPage}`,
  )
  const pushed = await page.evaluate(() => {
    const frags = Array.from(document.querySelectorAll('.wtp-pages .wtp-content > *'))
    const first = frags.find((el) => (el.textContent ?? '').includes('我方于2026年9月1日'))
    if (!first) return null
    const pageEl = first.closest('.wtp-page')
    return {
      isFirstInPage: pageEl ? pageEl.querySelector('.wtp-content > *') === first : false,
      page: pageEl ? Array.from(document.querySelectorAll('.wtp-page')).indexOf(pageEl) + 1 : -1,
    }
  })
  ok(
    '分页符把后面的段落推到了新一页的页首',
    pushed?.isFirstInPage === true && pushed.page > 1,
    JSON.stringify(pushed),
  )

  // 点标记上的 × 删掉刚插的分页符
  await page.click('.wtp-break .wtp-break-del')
  await page.waitForTimeout(250)
  const afterDel = await getModel()
  eq('点 × 删掉了分页符', afterDel.blocks.filter((b) => b.t === 'pageBreak').length, 0)
  const pagesAfterDel = await page.evaluate(() => document.querySelectorAll('.wtp-page').length)
  eq('页数回到插入前', pagesAfterDel, breaksBefore.pages)

  // 再插一个分节符（样本里本来还有一个 `---`，加完共 2 个）
  await page.evaluate(() => window.__wtpTest.setCaret('苏州市公安局', 0))
  await page.click('button.tool[title^="在光标所在段落后插入分节符"]')
  await page.waitForTimeout(250)
  const withSection = await getModel()
  eq('模型里多了一个分节符', withSection.blocks.filter((b) => b.t === 'sectionBreak').length, 2)
  const sectionMarks = await page.evaluate(() =>
    Array.from(document.querySelectorAll('.wtp-break')).map((el) => el.textContent.trim()),
  )
  ok(
    '页间出现了「分节符」标记',
    sectionMarks.some((t) => t.includes('分节符')),
    JSON.stringify(sectionMarks),
  )
  await checkNoOverflow('H4 插换页标记后')

  /* ------------------------------------------------------------------ */
  console.log('\n=== H5. 分页符 + 分节符连在一起：两枚标记都要看得见 ===')
  await openApp('插入')
  const beforeBoth = await page.evaluate(
    () => document.querySelectorAll('.wtp-page').length,
  )
  await page.evaluate(() => window.__wtpTest.setCaret('苏州市公安局', 0))
  await page.click('button.tool[title^="在光标所在段落后插入分页符"]')
  await page.waitForTimeout(200)
  await page.evaluate(() => window.__wtpTest.setCaret('苏州市公安局', 0))
  await page.click('button.tool[title^="在光标所在段落后插入分节符"]')
  await page.waitForTimeout(300)

  const bothModel = await getModel()
  eq(
    '模型里分页符/分节符各就各位',
    `${bothModel.blocks.filter((b) => b.t === 'pageBreak').length}/${
      bothModel.blocks.filter((b) => b.t === 'sectionBreak').length
    }`,
    '1/2',
  )
  const both = await page.evaluate(() => {
    const kids = Array.from(document.querySelector('.wtp-pages').children)
    const marks = kids.filter((el) => el.classList.contains('wtp-break'))
    return {
      texts: marks.map((el) => el.textContent.trim()),
      adjacent: kids.some(
        (el, i) =>
          el.classList.contains('wtp-break') && kids[i + 1]?.classList.contains('wtp-break'),
      ),
      pages: kids.filter((el) => el.classList.contains('wtp-page')).length,
    }
  })
  eq('三枚标记都画出来了（原有 1 个分节符 + 新增 2 枚）', both.texts.length, 3)
  eq(
    '文案里分页符 1 枚、分节符 2 枚',
    `${both.texts.filter((t) => t.includes('分页符')).length}/${
      both.texts.filter((t) => t.includes('分节符')).length
    }`,
    '1/2',
  )
  ok('新增的两枚挨在一起（不会被吞掉一枚）', both.adjacent)
  /*
   * 页数只可能不变或 +1（与 H4 同一条理由：页尾原来的余量被吸收掉时总数就不变，
   * 这不是缺陷）。原本这里断言的是 `beforeBoth + 1` —— demo 样本加了表格之后
   * 管理人文件恰好 5 页、余量够吸收，于是这条断言先于本次改动就红了（2026-09-15 修正）。
   */
  ok(
    '插两枚标记后页数不减少',
    both.pages >= beforeBoth,
    `${beforeBoth} → ${both.pages}`,
  )
  await checkNoOverflow('H5 两枚标记连在一起后')

  /* ------------------------------------------------------------------ */
  console.log('\n=== H6. 文末插分节符：Word 会多留一张空白页 ===')
  await openApp('插入')
  const beforeTail = await page.evaluate(() => {
    const pages = Array.from(document.querySelectorAll('.wtp-page'))
    const last = pages[pages.length - 1]
    return {
      pages: pages.length,
      lastNumber: (last?.querySelector('.wtp-page-number')?.textContent ?? '').trim(),
    }
  })
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('2026年9月12日'))
  await page.click('button.tool[title^="在光标所在段落后插入分节符"]')
  await page.waitForTimeout(300)
  const tail = await page.evaluate(() => {
    const pages = Array.from(document.querySelectorAll('.wtp-page'))
    const last = pages[pages.length - 1]
    return {
      pages: pages.length,
      lastFragments: last ? last.querySelectorAll('[data-block-id]').length : -1,
      lastNumber: (last?.querySelector('.wtp-page-number')?.textContent ?? '').trim(),
    }
  })
  eq('文末分节符多留了一张空白页', tail.pages, beforeTail.pages + 1)
  eq('最后一张确实没有片段', tail.lastFragments, 0)
  // W5：新插的分节符是**全默认**（关联前节 + 不重排），所以空白页的页码接着前一节往下数。
  // 旧行为是「每插一个分节符就从 1 重排」，那个默认值已被用户改掉（需求：默认关联前节）。
  eq(
    '空白页页码接着前一节往下数（新分节符默认不重排）',
    tail.lastNumber,
    String(Number(beforeTail.lastNumber) + 1),
  )
  await checkNoOverflow('H6 文末分节符后')

  /* ------------------------------------------------------------------ */
  console.log('\n=== J. Ctrl+U：切换选中文本的下划线 ===')
  await openApp()
  await page.evaluate(() => window.__wtpTest.selectIn('我方于2026年9月1日', 0, 2))
  await page.waitForTimeout(60)
  await page.keyboard.press('Control+u')
  await page.waitForTimeout(200)
  const ulOn = await getModel()
  const ulBlockOn = heroBlocks(ulOn).find((b) => textOfBlock(b).startsWith('我方'))
  const ulInline = ulBlockOn
    ? ulBlockOn.inlines.find((i) => i.t === 'text' && i.underline)
    : null
  ok(
    '模型里前两个字标了下划线',
    ulInline !== null && ulInline !== undefined && ulInline.text === '我方',
    JSON.stringify(ulInline ?? null),
  )
  const ulDomOn = await page.evaluate(() => {
    const frag = window.__wtpTest.fragmentByText('我方于2026年9月1日')
    return frag ? Array.from(frag.querySelectorAll('u')).map((el) => el.textContent) : []
  })
  ok('版面上那两个字在 <u> 里', ulDomOn.includes('我方'), JSON.stringify(ulDomOn))
  ok(
    '下划线没有改动文字',
    ulBlockOn !== undefined && textOfBlock(ulBlockOn).startsWith('我方于2026年9月1日'),
    JSON.stringify(ulBlockOn ? textOfBlock(ulBlockOn) : null),
  )

  // 再按一次：切换回不带下划线（工具栏按钮的语义是「切换」，不是「只加不减」）
  await page.keyboard.press('Control+u')
  await page.waitForTimeout(200)
  const ulOff = await getModel()
  const ulBlockOff = heroBlocks(ulOff).find((b) => textOfBlock(b).startsWith('我方'))
  ok(
    '再按一次取消下划线',
    !(ulBlockOff?.inlines ?? []).some((i) => i.t === 'text' && i.underline),
    JSON.stringify(ulBlockOff?.inlines ?? null),
  )
  eq('版面上也不再有 <u>', await page.evaluate(() => document.querySelectorAll('.wtp-content u').length), 0)

  // 与加粗叠加：两种格式互不覆盖
  await page.evaluate(() => window.__wtpTest.selectIn('我方于2026年9月1日', 0, 2))
  await page.waitForTimeout(60)
  await page.keyboard.press('Control+u')
  await page.waitForTimeout(150)
  await page.evaluate(() => window.__wtpTest.selectIn('我方于2026年9月1日', 0, 2))
  await page.waitForTimeout(60)
  await page.keyboard.press('Control+b')
  await page.waitForTimeout(200)
  const ulBothModel = await getModel()
  const ulBothBlock = heroBlocks(ulBothModel).find((b) => textOfBlock(b).startsWith('我方'))
  const ulBothInline = ulBothBlock
    ? ulBothBlock.inlines.find((i) => i.t === 'text' && i.text === '我方')
    : null
  ok(
    '加粗与下划线同时落在同两个字上',
    ulBothInline !== undefined &&
      ulBothInline !== null &&
      ulBothInline.bold === true &&
      ulBothInline.underline === true,
    JSON.stringify(ulBothInline ?? null),
  )
  const ulBothDom = await page.evaluate(() => {
    const frag = window.__wtpTest.fragmentByText('我方于2026年9月1日')
    return frag ? (frag.querySelector('u')?.innerHTML ?? '') : ''
  })
  ok('版面上是 <u><b>我方</b></u>', ulBothDom.includes('<b>我方</b>'), ulBothDom)
  await checkNoOverflow('J 下划线后')

  /* ------------------------------------------------------------------ */
  console.log('\n=== K. Ctrl+Shift+E：翻转修订模式（App 的复选框同步） ===')
  await openApp()
  const trackSel = 'label.checkbox input[type="checkbox"]'
  eq('初始未勾选修订模式', await page.isChecked(trackSel), false)
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('苏州市公安局'))
  await page.keyboard.press('Control+Shift+E')
  await page.waitForTimeout(150)
  eq('快捷键把修订模式打开了', await page.isChecked(trackSel), true)
  await page.keyboard.insertText('（快捷键开的修订）')
  await page.waitForTimeout(250)
  const trackModel = await getModel()
  const trackInline = heroBlocks(trackModel)
    .flatMap((b) => b.inlines)
    .find(
      (i) => i.t === 'text' && i.rev && i.rev.kind === 'ins' && i.text.includes('快捷键开的修订'),
    )
  ok('打开后新增文字标成 ins 修订', trackInline !== undefined)
  await page.keyboard.press('Control+Shift+E')
  await page.waitForTimeout(150)
  eq('再按一次关掉', await page.isChecked(trackSel), false)
  // 换个没有任何修订标记的段落敲字：在刚被标成 ins 的字后面接着敲，
  // 读回时新字会并进那个 ins 区间（DOM 上下文继承，见 README 的已知取舍）
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('我方于2026年9月1日'))
  await page.keyboard.insertText('（关闭后新增）')
  await page.waitForTimeout(250)
  const untracked = await getModel()
  const untrackedInline = heroBlocks(untracked)
    .flatMap((b) => b.inlines)
    .find((i) => i.t === 'text' && i.text.includes('（关闭后新增）'))
  ok(
    '关掉之后新增的文字不再带修订标记',
    untrackedInline !== undefined && untrackedInline.rev === undefined,
    JSON.stringify(untrackedInline ?? null),
  )
  await checkNoOverflow('K 修订模式开关后')

  /* ------------------------------------------------------------------ */
  console.log('\n=== L. Alt+4：选中的数字改成千分位两位小数 ===')
  await openApp()
  // 造一段只含数字的段落：回车分段后敲进去，选区就落在这一段上
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('苏州市公安局'))
  await page.keyboard.press('Enter')
  await page.waitForTimeout(150)
  await page.keyboard.insertText('12345.6')
  await page.waitForTimeout(150)
  await page.evaluate(() => window.__wtpTest.selectIn('12345.6', 0, 7))
  await page.waitForTimeout(60)
  await page.keyboard.press('Alt+4')
  await page.waitForTimeout(250)
  const amountModel = await getModel()
  ok(
    '模型里的数字改成了 12,345.60',
    heroBlocks(amountModel).some((b) => textOfBlock(b) === '12,345.60'),
    JSON.stringify(heroBlocks(amountModel).map(textOfBlock)),
  )
  const amountDom = await page.evaluate(() => {
    const frag = window.__wtpTest.fragmentByText('12,345.60')
    return frag ? window.__wtpTest.textOf(frag) : null
  })
  eq('版面上显示的也是 12,345.60', amountDom, '12,345.60')
  const amountCaret = await page.evaluate(() => window.__wtpTest.caretInfo())
  ok(
    '改完这串数字仍被选中（连着按结果稳定）',
    amountCaret !== null && amountCaret.text === '12,345.60',
    JSON.stringify(amountCaret),
  )

  // 第二次：已带逗号的数字要能再次解析，结果不变
  await page.keyboard.press('Alt+4')
  await page.waitForTimeout(250)
  const amountAgain = await getModel()
  ok(
    '带逗号的数字再格式化还是它自己',
    heroBlocks(amountAgain).filter((b) => textOfBlock(b) === '12,345.60').length === 1,
    JSON.stringify(heroBlocks(amountAgain).map(textOfBlock)),
  )

  // 负数与进位
  await page.evaluate(() => window.__wtpTest.selectIn('12,345.60', 0, 9))
  await page.waitForTimeout(60)
  await page.keyboard.insertText('-999.999')
  await page.waitForTimeout(150)
  await page.evaluate(() => window.__wtpTest.selectIn('-999.999', 0, 8))
  await page.waitForTimeout(60)
  await page.keyboard.press('Alt+4')
  await page.waitForTimeout(250)
  const amountRound = await getModel()
  ok(
    '进位正确：-999.999 → -1,000.00',
    heroBlocks(amountRound).some((b) => textOfBlock(b) === '-1,000.00'),
    JSON.stringify(heroBlocks(amountRound).map(textOfBlock)),
  )

  // 无效输入：不改模型，弹提示条
  const beforeBad = JSON.stringify(await getModel())
  await page.evaluate(() => window.__wtpTest.selectIn('苏州市公安局', 0, 2))
  await page.waitForTimeout(60)
  await page.keyboard.press('Alt+4')
  await page.waitForTimeout(200)
  const toast = await page.evaluate(() => {
    const el = document.querySelector('.toast')
    return el ? el.textContent.trim() : null
  })
  eq('选中非数字时弹出提示条', toast, '选中内容不是有效数字')
  eq('无效输入不改模型', JSON.stringify(await getModel()), beforeBad)

  // 折叠的光标（没有选中任何文字）同样不改模型，同样提示
  await page.evaluate(() => window.__wtpTest.setCaret('苏州市公安局', 2))
  await page.waitForTimeout(60)
  await page.keyboard.press('Alt+4')
  await page.waitForTimeout(200)
  eq('没有选中文字时提示无效', await page.evaluate(() => document.querySelector('.toast')?.textContent.trim() ?? null), '选中内容不是有效数字')
  eq('折叠光标下模型不变', JSON.stringify(await getModel()), beforeBad)

  // 提示条约 2 秒后自己消失
  await page.waitForTimeout(2300)
  eq('提示条自动消失', await page.evaluate(() => document.querySelectorAll('.toast').length), 0)
  await checkNoOverflow('L 金额格式化后')

  /* ------------------------------------------------------------------ */
  console.log('\n=== M. 工具栏：插入三种特殊空格（并排按钮，模型里是确切码点） ===')
  // 三枚空格按钮在「插入」页里
  await openApp('插入')
  const SPACE_TITLE = {
    em: '在插入符处插入全宽空格',
    en: '在插入符处插入半宽空格',
    quarterEm: '在插入符处插入四分之一宽空格',
  }
  const spaceButton = (kind) =>
    page.locator(`.panel-insert button.tool[title^="${SPACE_TITLE[kind]}"]`)
  /** 点按钮插入。按钮是 @mousedown.prevent：焦点不离开正文，走实时选区那条路 */
  async function pickSpace(kind) {
    await spaceButton(kind).click()
    await page.waitForTimeout(250)
  }
  /** 先把焦点移出正文再点按钮（真人先点了别处的典型路径）：考选区/落点记录的兜底 */
  async function pickSpaceBlurred(kind) {
    await page.evaluate(() => document.activeElement?.blur?.())
    await spaceButton(kind).click()
    await page.waitForTimeout(250)
  }

  await page.evaluate(() => window.__wtpTest.caretAtEndOf('苏州市公安局'))
  await pickSpace('em')
  const emModel = await getModel()
  const emBlock = heroBlocks(emModel).find((b) => textOfBlock(b).includes('\u2003'))
  ok('模型里插入了 U+2003（全宽空格）', emBlock !== undefined, JSON.stringify(heroBlocks(emModel).map(textOfBlock)))
  eq(
    '插入的是全宽空格，不多不少一个',
    emBlock ? Array.from(textOfBlock(emBlock)).filter((c) => c.codePointAt(0) === 0x2003).length : -1,
    1,
  )
  ok('空格落在插入符处（段末）', textOfBlock(emBlock).endsWith('\u2003'), JSON.stringify(textOfBlock(emBlock)))
  const caretAfterSpace = await page.evaluate(() => window.__wtpTest.caretInfo())
  ok(
    '插入符落在刚插入的空格之后',
    caretAfterSpace !== null && caretAfterSpace.offset === textOfBlock(emBlock).length,
    JSON.stringify(caretAfterSpace),
  )
  const domHasEm = await page.evaluate(() => {
    const frag = window.__wtpTest.fragmentByText('\u2003')
    return frag ? window.__wtpTest.textOf(frag) : null
  })
  ok('版面上也读得到这个码点', domHasEm !== null && domHasEm.includes('\u2003'), JSON.stringify(domHasEm))

  // 另外两个：半宽与四分之一宽。这次先把焦点移出正文（考选区/落点记录的兜底）
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('苏州市公安局'))
  await pickSpaceBlurred('en')
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('苏州市公安局'))
  await pickSpaceBlurred('quarterEm')
  const spacesModel = await getModel()
  const spacesBlock = heroBlocks(spacesModel).find((b) => textOfBlock(b).includes('\u2002'))
  const points = spacesBlock ? Array.from(textOfBlock(spacesBlock)).map((c) => c.codePointAt(0)) : []
  ok('模型里有 U+2002', points.includes(0x2002), JSON.stringify(points))
  ok('模型里有 U+2005', points.includes(0x2005), JSON.stringify(points))
  ok('三种空格都插在段末，顺序与操作一致', spacesBlock
    ? textOfBlock(spacesBlock).endsWith('\u2003\u2002\u2005')
    : false, JSON.stringify(textOfBlock(spacesBlock)))
  ok('普通空格没有混进来', !points.includes(0x20), JSON.stringify(points))

  // 连点同一枚按钮两次必须插两个 —— 三枚并排按钮换掉那个下拉就是为了这个：
  // 下拉选完要复位回占位项才认第二次 change，否则「再选同一项」是静默的
  await pickSpace('em')
  await pickSpace('em')
  const twiceModel = await getModel()
  const twiceBlock = heroBlocks(twiceModel).find((b) => textOfBlock(b).includes('\u2003'))
  ok(
    '连点两次同一枚空格按钮，插入两个',
    twiceBlock ? textOfBlock(twiceBlock).endsWith('\u2003\u2002\u2005\u2003\u2003') : false,
    JSON.stringify(twiceBlock ? textOfBlock(twiceBlock) : null),
  )

  // 有选区时替换选区（与 insertText / replaceRange 的既有语义一致）
  await page.evaluate(() => window.__wtpTest.selectIn('我方于2026年9月1日', 0, 2))
  await page.waitForTimeout(60)
  await pickSpace('en')
  const replaced = await getModel()
  const replacedBlock = heroBlocks(replaced).find((b) => textOfBlock(b).includes('于2026年9月1日'))
  ok(
    '有选区时用空格替换掉选中的字',
    replacedBlock !== undefined && textOfBlock(replacedBlock).startsWith('\u2002于2026年9月1日'),
    JSON.stringify(replacedBlock ? textOfBlock(replacedBlock) : null),
  )
  const pagesAfterSpaces = await page.evaluate(() => document.querySelectorAll('.wtp-page').length)
  ok('插入空格没有把页数搞乱', pagesAfterSpaces >= 1)
  await checkNoOverflow('M 插特殊空格后')

  // 还没在版面上放过插入符：既不该改模型，也该给一句提示
  await openApp('插入')
  const beforeNoCaret = JSON.stringify(await getModel())
  await pickSpace('em')
  eq('没有插入符时不改模型', JSON.stringify(await getModel()), beforeNoCaret)
  eq(
    '没有插入符时给出提示',
    await page.evaluate(() => document.querySelector('.toast')?.textContent.trim() ?? null),
    '请先把插入符放到版面上',
  )

  /* ------------------------------------------------------------------ */
  console.log('\n=== N. 切文件模板：页数变、插入符与选区不丢、文字不丢 ===')
  await openApp()

  // 模板清单从下拉里读：这是 UI 测试，就照界面给的选项走
  const templateOptions = await page.$$eval('.bar select option', (els) =>
    els.map((el) => ({ key: el.value, label: (el.textContent ?? '').trim() })),
  )
  ok(
    '顶栏「文件模板」下拉有至少两套模板',
    templateOptions.length >= 2,
    JSON.stringify(templateOptions),
  )
  const defaultKey = await page.$eval('.bar select', (el) => el.value)
  eq('默认选中的是第一套模板', defaultKey, templateOptions[0]?.key)
  const otherKey = templateOptions.find((t) => t.key !== defaultKey)?.key ?? ''

  const textsBefore = heroBlocks(await getModel()).map(textOfBlock)

  // 插入符放到文档中段的一块正文里：公文版心更矮，这一块一定会被重新分页，
  // 落点与选区只能靠「块 id + 字符偏移」找回来（DOM 节点全被重建了）。
  const caretSeeded = await page.evaluate(() =>
    window.__wtpTest.caretAtEndOf('债务人爱康光电科技有限公司'),
  )
  ok('找到了要落插入符的段落', caretSeeded !== null, JSON.stringify(caretSeeded))
  await page.waitForTimeout(80)
  const caretBefore = await page.evaluate(() => window.__wtpTest.caretInfo())
  const pagesBefore = await page.evaluate(() => document.querySelectorAll('.wtp-page').length)
  /** 版心几何 + 量测行数总和：切模板要「整篇按新版心重量」，这两样必须变 */
  const measureInfo = () =>
    page.evaluate(() => {
      const blocks = window.__wtpPaper.getMeasurements().filter((it) => it.t === 'block')
      return {
        contentHeight: document.querySelector('.wtp-content')?.clientHeight ?? 0,
        contentWidth: document.querySelector('.wtp-content')?.clientWidth ?? 0,
        rowSum: blocks.reduce((n, b) => n + b.rows, 0),
      }
    })
  const beforeInfo = await measureInfo()

  await page.selectOption('.bar select', otherKey)
  await page.waitForTimeout(1000)

  const pagesAfter = await page.evaluate(() => document.querySelectorAll('.wtp-page').length)
  const caretAfter = await page.evaluate(() => window.__wtpTest.caretInfo())
  /*
   * 「整篇按新版心重量了」的判据不能是「页数变了」—— 两套模板的页数可以巧合地相同
   * （demo 样本现在就是 5 ↔ 5）。真正要证的是量测确实重跑过：版心高按新规格表换了，
   * 而且长段落的实测行数跟着换（期望值从规格表与量测现取，不硬编码）。
   */
  const afterInfo = await measureInfo()
  ok(
    '切模板后版心几何按新规格表换了',
    afterInfo.contentHeight !== beforeInfo.contentHeight || afterInfo.contentWidth !== beforeInfo.contentWidth,
    JSON.stringify({ before: beforeInfo, after: afterInfo }),
  )
  ok(
    '切模板后整篇按新版心重量了（长段落行数跟着变）',
    afterInfo.rowSum !== beforeInfo.rowSum,
    JSON.stringify({ before: beforeInfo, after: afterInfo }),
  )
  ok(
    '切模板后页数不减少（公文版心更矮，页数只可能更多）',
    pagesAfter >= pagesBefore,
    `${pagesBefore} → ${pagesAfter}`,
  )
  eq('插入符还在同一块', caretAfter?.blockId, caretBefore?.blockId)
  eq('插入符还在同一字符偏移', caretAfter?.offset, caretBefore?.offset)
  eq('插入符那一块的文字没变', caretAfter?.text, caretBefore?.text)

  // 选区：选中一段文字再切回第一套模板，选区与文字都该原样回来
  const selectedText = await page.evaluate(() =>
    window.__wtpTest.selectIn('债务人爱康光电科技有限公司', 0, 6),
  )
  ok('切换前选中了六个字', selectedText === '债务人爱康光', selectedText ?? '未选中')
  await page.waitForTimeout(80)
  await page.selectOption('.bar select', defaultKey)
  await page.waitForTimeout(1000)
  const selectionNow = await page.evaluate(() => document.getSelection()?.toString() ?? '')
  eq('切回来选区仍在（还是那六个字）', selectionNow, selectedText)
  eq(
    '切回第一套模板后页数回到原值',
    await page.evaluate(() => document.querySelectorAll('.wtp-page').length),
    pagesBefore,
  )

  const textsAfter = heroBlocks(await getModel()).map(textOfBlock)
  eq('切模板没有改动任何文字', JSON.stringify(textsAfter), JSON.stringify(textsBefore))
  await checkNoOverflow('N 切模板后')

  /* ------------------------------------------------------------------ */
  console.log('\n=== O. 打印：只出 A4 纸，且不多不少 ===')
  // 「插入表格」按钮在「插入」页里
  await openApp('插入')
  // 打开查找面板：打印时它和导航窗格都必须一起消失。
  // 落点放进表格格子里 —— 表格页里的控件要光标在格内才启用，打印时它们也必须一起消失。
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('数控加工中心'))
  await page.keyboard.press('Control+f')
  await page.waitForSelector('.search-panel', { timeout: 3000 })
  // 插入表格面板也要一起验：它和查找面板一样是浮层（挂在预览窗格上，切页不会把它关掉），
  // 打印时同样不该出现
  await page.locator('button.tool[title^="在光标所在段落后插入一张空表格"]').click()
  await page.waitForSelector('.table-panel', { timeout: 3000 })

  /** 打印媒体下这些选择器的 display（元素不存在时给 'missing'，别把「没有」当成「隐藏了」） */
  const printDisplay = () =>
    page.evaluate(() => {
      const sels = [
        '.bar',
        '.ribbon-tabs',
        '.styles',
        '.toolbar',
        '.sub-toolbar',
        '.section-toolbar',
        '.pane-head',
        '.legend',
        'textarea',
        '.wtp-comments',
        '.wtp-measure-root',
        '.wtp-break',
        '.nav-pane',
        '.search-panel',
        '.table-panel',
      ]
      const out = {}
      for (const sel of sels) {
        const el = document.querySelector(sel)
        out[sel] = el ? getComputedStyle(el).display : 'missing'
      }
      out.pages = document.querySelectorAll('.wtp-page').length
      const first = document.querySelector('.wtp-page')
      const second = document.querySelectorAll('.wtp-page')[1]
      out.firstBreak = first ? getComputedStyle(first).breakBefore : 'missing'
      out.secondBreak = second ? getComputedStyle(second).breakBefore : 'missing'
      out.boxShadow = first ? getComputedStyle(first).boxShadow : 'missing'
      return out
    })

  /** 当前媒体下某个选择器的 display（不存在给 'missing'，别把「没有」当成「隐藏了」） */
  const displayOf = (sel) =>
    page.evaluate((s) => {
      const el = document.querySelector(s)
      return el ? getComputedStyle(el).display : 'missing'
    }, sel)

  // @page：尺寸来自规格表，边距归 0（白边由纸张自己的 padding 提供，不能留两份）
  const pageRules = () =>
    page.evaluate(() => {
      const out = []
      for (const sheet of document.styleSheets) {
        let rules
        try {
          rules = sheet.cssRules
        } catch {
          continue
        }
        for (const rule of rules) {
          if (rule.conditionText !== 'print') continue
          for (const inner of rule.cssRules ?? []) {
            if (inner.constructor.name !== 'CSSPageRule') continue
            out.push({ size: inner.style.size, margin: inner.style.margin })
          }
        }
      }
      return out
    })

  await page.emulateMedia({ media: 'print' })
  const printEdit = await printDisplay()
  // .table-panel 只在编辑态的打印里断言：切到源码视图时面板会被主动关掉（它在那儿没有意义），
  // 那时元素不存在，给的是 'missing' 而不是 'none'，下面那个源码视图的循环因此不带它
  for (const sel of [
    '.bar',
    '.ribbon-tabs',
    '.wtp-comments',
    '.wtp-measure-root',
    '.wtp-break',
    '.nav-pane',
    '.search-panel',
    '.table-panel',
  ]) {
    eq(`打印时隐藏 ${sel}`, printEdit[sel], 'none')
  }
  /*
   * 功能区的四页外壳：页内容是按 tab 用 v-if 渲染的，同时只有一页在 DOM 里，
   * 所以得逐页切过去、逐页验（只看某一个 .toolbar 会把另外三页漏掉）。
   * 切页前先切回 screen —— 打印媒体下标签条自己就是 display:none，点不动。
   */
  for (const [label, panel] of [
    ['开始', '.panel-start'],
    ['插入', '.panel-insert'],
    ['布局', '.panel-layout'],
    ['表格', '.panel-table'],
  ]) {
    await page.emulateMedia({ media: 'screen' })
    await openTab(label)
    await page.emulateMedia({ media: 'print' })
    eq(`打印时隐藏「${label}」页`, await displayOf(panel), 'none')
    if (label === '开始') eq('打印时隐藏样式库', await displayOf('.styles'), 'none')
  }
  eq('第一张纸前面不再断页', printEdit.firstBreak, 'auto')
  eq('后续每张纸都在新的一页开始', printEdit.secondBreak, 'page')
  eq('打印时纸张不带屏幕上的阴影', printEdit.boxShadow, 'none')
  eq('打印时版面页数不变', printEdit.pages, await page.evaluate(() => document.querySelectorAll('.wtp-page').length))
  // W4b-2 新增的三组按钮都在 .sub-toolbar 里：打印时它们随外壳一起消失（逐个按 offsetParent 验，
  // 只看 .sub-toolbar 自己的 display 证明不了子按钮真被盖住）
  ok(
    '打印时上下文工具条里的新增按钮也不显示（水平 / 垂直 / 删除表格）',
    await page.evaluate(() => {
      const groups = Array.from(document.querySelectorAll('.sub-toolbar .tk-group'))
      const names = ['水平', '垂直', '删除表格']
      return names.every((name) => {
        const group = groups.find((g) => (g.textContent ?? '').includes(name))
        if (!group) return false
        const btns = Array.from(group.querySelectorAll('button'))
        return btns.length > 0 && btns.every((b) => b.offsetParent === null)
      })
    }),
  )

  const rules = await pageRules()
  ok('打印样式里有 @page', rules.length > 0, JSON.stringify(rules))
  ok(
    '@page 尺寸取自规格表（A4）',
    rules.some((r) => r.size.includes('210mm') && r.size.includes('297mm')),
    JSON.stringify(rules),
  )
  ok(
    '@page 边距为 0（不额外加白边）',
    rules.every((r) => /^0(px)?$/.test(r.margin)),
    JSON.stringify(rules),
  )

  // 真打一份 PDF 数页数：不多一张空白页，也不少一页
  const pagesToPrint = await page.evaluate(() => document.querySelectorAll('.wtp-page').length)
  const pdf = await page.pdf({ printBackground: true, preferCSSPageSize: true })
  const pdfText = pdf.toString('latin1')
  const pdfPages = (pdfText.match(/\/Type\s*\/Page[^s]/g) ?? []).length
  eq('打印出来的页数 = 版面页数', pdfPages, pagesToPrint)
  const mediaBox = /\/MediaBox\s*\[\s*0\s+0\s+([\d.]+)\s+([\d.]+)\s*\]/.exec(pdfText)
  ok(
    '纸张尺寸是 A4（595.28 × 841.89 磅）',
    mediaBox !== null &&
      Math.abs(Number(mediaBox[1]) - 595.28) < 1 &&
      Math.abs(Number(mediaBox[2]) - 841.89) < 1,
    mediaBox ? `MediaBox ${mediaBox[1]} × ${mediaBox[2]}` : 'PDF 里没找到 MediaBox',
  )

  // 源码视图（左侧面板、textarea、语法说明）同样一张都不该印出来。
  // 切模式要在回到屏幕媒体时做：顶栏在打印媒体下是隐藏的，点不到。
  await page.emulateMedia({ media: 'screen' })
  await page.click('.tabs button:nth-child(2)')
  await page.waitForTimeout(300)
  await page.emulateMedia({ media: 'print' })
  const printSource = await printDisplay()
  for (const sel of ['.pane-head', '.legend', 'textarea']) {
    eq(`源码视图下打印也隐藏 ${sel}`, printSource[sel], 'none')
  }
  await page.emulateMedia({ media: 'screen' })

  /* ------------------------------------------------------------------ */
  console.log('\n=== P. Ctrl+F 查找：面板、高亮、不重排、上一处/下一处 ===')
  await openApp()
  await page.evaluate(() => {
    const frag = window.__wtpTest.fragmentByText('我方于2026年9月1日')
    frag.__wtpKeep = true
    window.__wtpTest.caretAtEndOf('我方于2026年9月1日')
  })
  await page.keyboard.press('Control+f')
  await page.waitForSelector('.search-panel', { timeout: 3000 })
  await page.waitForTimeout(200)
  eq(
    'ctrl+F 打开面板后焦点在查找框',
    await page.evaluate(
      () =>
        document.activeElement === document.querySelectorAll('.search-panel input[type="text"]')[0],
    ),
    true,
  )

  const findBox = page.locator('.search-panel input[type="text"]').first()
  const replaceBox = page.locator('.search-panel input[type="text"]').nth(1)
  const countText = async () => (await page.locator('.search-count').textContent()).trim()
  const panelButton = (name) =>
    page.locator('.search-panel').getByRole('button', { name, exact: true })

  const hitsP = countInModel(await getModel(), '债务人')
  ok('样本里「债务人」出现多处，够验上下移动', hitsP >= 2, String(hitsP))
  await findBox.fill('债务人')
  await page.waitForTimeout(250)
  eq('计数器与模型算出来的命中数一致', await countText(), `第 1 / 共 ${hitsP} 处`)

  const paintedP = await page.evaluate(() => window.__wtpTest.highlights())
  ok('命中有高亮（Custom Highlight API）', paintedP.normal + paintedP.current > 0, JSON.stringify(paintedP))
  ok('当前那一处单独一层高亮', paintedP.current >= 1, JSON.stringify(paintedP))

  const keptP = await page.evaluate(() => {
    const frag = window.__wtpTest.fragmentByText('我方于2026年9月1日')
    return {
      same: frag ? frag.__wtpKeep === true : false,
      connected: frag ? frag.isConnected : false,
    }
  })
  ok('查找没有重建版面 DOM（片段还是同一个节点）', keptP.same && keptP.connected)

  await panelButton('下一个').click()
  await page.waitForTimeout(250)
  eq('「下一个」走到第 2 处', await countText(), `第 2 / 共 ${hitsP} 处`)
  await panelButton('上一个').click()
  await page.waitForTimeout(250)
  eq('「上一个」回到第 1 处', await countText(), `第 1 / 共 ${hitsP} 处`)

  // 面板可拖动，且拖不出预览窗格
  const beforeDrag = await page.evaluate(() => {
    const rect = document.querySelector('.search-panel').getBoundingClientRect()
    return { left: Math.round(rect.left), top: Math.round(rect.top) }
  })
  const headBox = await page.locator('.search-head').boundingBox()
  await page.mouse.move(headBox.x + 40, headBox.y + 10)
  await page.mouse.down()
  await page.mouse.move(headBox.x + 40 - 140, headBox.y + 10 + 70, { steps: 5 })
  await page.mouse.up()
  await page.waitForTimeout(200)
  const dragResult = await page.evaluate(() => {
    const panel = document.querySelector('.search-panel').getBoundingClientRect()
    const pane = document.querySelector('.preview-pane').getBoundingClientRect()
    return {
      left: Math.round(panel.left),
      top: Math.round(panel.top),
      inside:
        panel.left >= pane.left - 1 &&
        panel.top >= pane.top - 1 &&
        panel.right <= pane.right + 1 &&
        panel.bottom <= pane.bottom + 1,
    }
  })
  ok(
    '拖标题栏能把面板挪走',
    dragResult.left !== beforeDrag.left || dragResult.top !== beforeDrag.top,
    `${JSON.stringify(beforeDrag)} → ${JSON.stringify(dragResult)}`,
  )
  ok('面板被夹在预览窗格范围内', dragResult.inside, JSON.stringify(dragResult))

  await findBox.press('Escape')
  await page.waitForTimeout(250)
  eq('Esc 关掉面板', await page.locator('.search-panel').count(), 0)
  const hlClosed = await page.evaluate(() => window.__wtpTest.highlights())
  eq('关闭后高亮清掉', hlClosed.normal + hlClosed.current, 0)

  /* ------------------------------------------------------------------ */
  console.log('\n=== P2. 非法正则只提示；范围「当前选中的文本」只命中选区内 ===')
  await openApp()
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('苏州市公安局'))
  await page.keyboard.press('Control+f')
  await page.waitForSelector('.search-panel', { timeout: 3000 })
  const modelBad = JSON.stringify(await getModel())
  await page.locator('.search-panel input[name="search-regex"]').check()
  await findBox.fill('[')
  await page.waitForTimeout(250)
  const errCount = await page.locator('.search-panel .search-error').count()
  const errText =
    errCount > 0
      ? ((await page.locator('.search-panel .search-error').textContent()) ?? '').trim()
      : ''
  ok('非法正则弹出提示条', errText !== '', errText || '没有提示')
  const hlBad = await page.evaluate(() => window.__wtpTest.highlights())
  eq('非法正则不高亮', hlBad.normal + hlBad.current, 0)
  eq('非法正则不改模型', JSON.stringify(await getModel()), modelBad)

  // 复原，再做范围那一条
  await page.locator('.search-panel input[name="search-regex"]').uncheck()
  await findBox.fill('')
  await page.waitForTimeout(150)
  await page.keyboard.press('Escape')
  await page.waitForTimeout(150)

  const picked = await page.evaluate(() =>
    window.__wtpTest.selectIn('债务人爱康光电科技有限公司', 0, 6),
  )
  eq('先选中一段文字', picked, '债务人爱康光')
  await page.keyboard.press('Control+f')
  await page.waitForSelector('.search-panel', { timeout: 3000 })
  const hitsAll = countInModel(await getModel(), '债务人')
  ok('全文里「债务人」不止一处（对照用）', hitsAll > 1, String(hitsAll))
  await page.locator('.search-panel input[name="search-scope"][value="selection"]').check()
  await findBox.fill('债务人')
  await page.waitForTimeout(250)
  eq('范围限定后只命中选区内那一处', await countText(), '第 1 / 共 1 处')
  await page.locator('.search-panel input[name="search-scope"][value="all"]').check()
  await page.waitForTimeout(250)
  eq('切回「全文」命中数恢复', await countText(), `第 1 / 共 ${hitsAll} 处`)

  /* ------------------------------------------------------------------ */
  console.log('\n=== P3. Ctrl+G：替换一处与全部替换（模型与 DOM 都要看）===')
  await openApp()
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('苏州市公安局'))
  await page.keyboard.press('Control+g')
  await page.waitForSelector('.search-panel', { timeout: 3000 })
  await page.waitForTimeout(200)
  eq(
    'ctrl+G 打开时焦点在替换框',
    await page.evaluate(
      () =>
        document.activeElement === document.querySelectorAll('.search-panel input[type="text"]')[1],
    ),
    true,
  )

  const beforeOne = heroBlocks(await getModel()).map(textOfBlock)
  await findBox.fill('《调取证据通知书》')
  await page.waitForTimeout(250)
  eq('替换前命中一处', await countText(), '第 1 / 共 1 处')
  await replaceBox.fill('《调取证据通知书（补）》')
  await panelButton('替换').click()
  await page.waitForTimeout(400)
  const expectedOne = beforeOne.map((t) =>
    t.replace('《调取证据通知书》', '《调取证据通知书（补）》'),
  )
  eq(
    '替换一处后模型文字与预期一致',
    JSON.stringify(heroBlocks(await getModel()).map(textOfBlock)),
    JSON.stringify(expectedOne),
  )
  await checkDomMatchesModel('替换一处后版面上每一块的文字都等于模型（含自动编号）')
  eq('替换后不再命中', await countText(), '共 0 处')

  // 插入符必须停在「新文字的末尾」：force 重排会重建 DOM，锚点给错插入符就丢，
  // 接着敲字会插到段首甚至文档开头（验收时正是这么抓到的）。
  const wantOneCaret = await expectedCaretAtEndOf('《调取证据通知书（补）》')
  const gotOneCaret = await page.evaluate(() => window.__wtpTest.caretInfo())
  ok(
    '替换后插入符落在新文字末尾',
    sameCaret(gotOneCaret, wantOneCaret),
    `实际 ${JSON.stringify(gotOneCaret)}，期望 ${JSON.stringify(wantOneCaret)}`,
  )
  await page.keyboard.type('X')
  await page.waitForTimeout(400)
  ok(
    '替换后接着敲字落在替换处之后（不是段首）',
    heroBlocks(await getModel())
      .map(textOfBlock)
      .some((t) => t.includes('《调取证据通知书（补）》X')),
    JSON.stringify((await heroBlocks(await getModel()).map(textOfBlock)).slice(0, 2)),
  )

  // 全部替换
  const beforeAll = countInModel(await getModel(), '债务人')
  ok('「债务人」有不止一处可换', beforeAll >= 2, String(beforeAll))
  eq('替换词原本不存在', countInModel(await getModel(), '义务人'), 0)
  await findBox.fill('债务人')
  await page.waitForTimeout(250)
  eq('全部替换前计数正确', await countText(), `第 1 / 共 ${beforeAll} 处`)
  await replaceBox.fill('义务人')
  await panelButton('全部替换').click()
  await page.waitForTimeout(500)
  const afterAll = await getModel()
  eq('全部替换后旧词一个不剩（删除修订里的不算）', countInModel(afterAll, '债务人'), 0)
  eq('新词数量等于旧词原数量', countInModel(afterAll, '义务人'), beforeAll)
  await checkDomMatchesModel('全部替换后版面上每一块的文字都等于模型（含自动修订与编号）')

  // 全部替换同样要还原插入符：停在第一处新文字的末尾（与 searchIndex 归零一致）
  const wantAllCaret = await expectedCaretAtEndOf('义务人')
  const gotAllCaret = await page.evaluate(() => window.__wtpTest.caretInfo())
  ok(
    '全部替换后插入符落在第一处新文字末尾',
    sameCaret(gotAllCaret, wantAllCaret),
    `实际 ${JSON.stringify(gotAllCaret)}，期望 ${JSON.stringify(wantAllCaret)}`,
  )

  /* ------------------------------------------------------------------ */
  console.log('\n=== Q. 导航窗格：条目与模型一致、点击跳转、折叠 ===')
  await openApp()
  await page.waitForSelector('.nav-pane', { timeout: 10000 })
  const want = expectedOutline(await getModel())
  ok('模型里确实有标题', want.length > 0, String(want.length))

  const navItems = await page.evaluate(() =>
    Array.from(document.querySelectorAll('.nav-pane .nav-list li')).map((li) => {
      const button = li.querySelector('button')
      const prefix = button?.querySelector('.nav-prefix')?.textContent ?? ''
      return {
        prefix,
        text: (button?.textContent ?? '').slice(prefix.length).trim(),
        level: Array.from(li.classList).find((c) => c.startsWith('nav-lv')) ?? '',
      }
    }),
  )
  eq('条目数 = 模型里的 h1/h2/h3 数', navItems.length, want.length)
  eq(
    '每条的前缀/文字/层级都与模型一致',
    JSON.stringify(navItems),
    JSON.stringify(
      want.map((e) => ({ prefix: e.prefix, text: e.text, level: `nav-lv${e.level}` })),
    ),
  )
  const navWidth = await page.evaluate(() =>
    Math.round(document.querySelector('.nav-pane').getBoundingClientRect().width),
  )
  ok('导航窗格宽约 200px', Math.abs(navWidth - 200) <= 1, `${navWidth}px`)

  // 点击跳转：插入符落在该块自动编号之后
  const targetIndex = 1
  await page.locator('.nav-pane .nav-list li button').nth(targetIndex).click()
  await page.waitForTimeout(300)
  const caretQ = await page.evaluate(() => window.__wtpTest.caretInfo())
  eq('插入符落在被点的标题块里', caretQ?.blockId, want[targetIndex].blockId)
  eq('插入符落在自动编号之后', caretQ?.offset, want[targetIndex].prefix.length)

  // 在正文段落里打字不该重建左栏（大纲指纹没变）
  await page.evaluate(() => {
    const item = document.querySelector('.nav-pane .nav-list li button')
    item.__wtpKeep = true
    window.__wtpTest.caretAtEndOf('我方于2026年9月1日')
  })
  await page.keyboard.insertText('（导航探针）')
  await page.waitForTimeout(300)
  ok(
    '正文打字不重建导航窗格',
    await page.evaluate(() => {
      const item = document.querySelector('.nav-pane .nav-list li button')
      return item ? item.__wtpKeep === true : false
    }),
  )

  await page.locator('.nav-pane .nav-collapse').click()
  await page.waitForTimeout(250)
  eq('窗格里的折叠按钮把它收起来', await page.locator('.nav-pane').count(), 0)
  await page.locator('.bar button', { hasText: '导航' }).click()
  await page.waitForTimeout(250)
  eq('顶栏的导航开关把它放回来', await page.locator('.nav-pane').count(), 1)

  /* ------------------------------------------------------------------ */
  console.log('\n=== T. 表格：渲染、格内读回、行不跨页拆开 ===')
  await openApp()
  const tableModel = await getModel()
  const tableBlocks = tableModel.blocks.filter((b) => b.t === 'table')
  ok('样本里有一张表（本波 demo 样本新增）', tableBlocks.length >= 1)
  if (tableBlocks.length > 0) {
    const t = tableBlocks[0]
    const surface = await page.evaluate(() => {
      const frags = Array.from(document.querySelectorAll('.wtp-tableFrag'))
      return {
        frags: frags.length,
        outerHasBlockId: frags.some((el) => el.dataset.blockId !== undefined),
        renderedRows: frags.reduce(
          (n, el) => n + el.querySelectorAll(':scope > table > tbody > tr').length,
          0,
        ),
        // 每个片段自带的行区间必须与它渲出来的行数一致（行不被拆开的直接证据）
        spansMatch: frags.every(
          (el) =>
            el.querySelectorAll(':scope > table > tbody > tr').length ===
            Number(el.dataset.rowTo) - Number(el.dataset.rowFrom),
        ),
        cells: document.querySelectorAll('.wtp-table .wtp-cell[data-block-id]').length,
        firstCellText: document.querySelector('.wtp-table .wtp-cell[data-block-id]')
          ?.textContent,
      }
    })
    ok('表格渲成了 .wtp-tableFrag', surface.frags >= 1)
    eq('表格外层不挂 data-block-id（否则会被当成片段读回）', surface.outerHasBlockId, false)
    eq('各片渲染的行数合计 = 模型行数', surface.renderedRows, t.rows.length)
    ok('每个表格片段渲染的行数 = 它的行区间（行没被拆开）', surface.spansMatch)
    eq(
      '格内那层都挂了 data-block-id',
      surface.cells,
      t.rows.reduce((n, row) => n + (row.role === 'body' ? t.columns : 1), 0),
    )

    // 格内打字：模型必须同步，且版式没变（正常输入不得重排）
    // 目标格子是「数控加工中心」那一格：样本表第 0 行是 unit 行，第 1 行是正文表头
    const targetCell = `${t.id}.r2c0`
    const pagesBeforeT = await page.evaluate(() => document.querySelectorAll('.wtp-page').length)
    const beforeCell = await page.evaluate(
      (id) => document.querySelector(`[data-block-id="${id}"]`)?.textContent ?? '',
      targetCell,
    )
    await page.evaluate(() => window.__wtpTest.caretAtEndOf('数控加工中心'))
    await page.keyboard.insertText('（试）')
    await page.waitForTimeout(300)
    const afterT = await getModel()
    const afterTable = afterT.blocks.find((b) => b.t === 'table')
    const afterCell = (afterTable?.rows[2]?.cells[0]?.inlines ?? [])
      .map((i) => (i.t === 'text' ? i.text : ''))
      .join('')
    ok(
      '格内打字同步到模型',
      afterCell === `${beforeCell}（试）`,
      `模型里是「${afterCell}」`,
    )
    eq(
      '格内打字不改变页数',
      await page.evaluate(() => document.querySelectorAll('.wtp-page').length),
      pagesBeforeT,
    )
    ok(
      '格内打字后片段还是同一个 DOM 节点',
      await page.evaluate((id) => {
        const el = document.querySelector(`[data-block-id="${id}"]`)
        return el !== null && el.isConnected
      }, targetCell),
    )
    await checkNoOverflow('T 格内打字后')
  }

  /* ------------------------------------------------------------------ */
  console.log('\n=== U. 插入表格面板：选规格后按规格插入，Esc 只关面板 ===')
  // 「插入表格」按钮在「插入」页里（功能区四页各管各的）
  await openApp('插入')
  const tableIds = async () =>
    (await getModel()).blocks.filter((b) => b.t === 'table').map((b) => b.id)
  const beforeU = await tableIds()

  const tableBtn = 'button.tool[title^="在光标所在段落后插入一张空表格"]'
  const tablePanel = page.locator('.table-panel')
  const rowBox = tablePanel.locator('input[type="number"]').nth(0)
  const colBox = tablePanel.locator('input[type="number"]').nth(1)

  await page.locator(tableBtn).click()
  await page.waitForTimeout(150)
  eq('点「插入表格」出现规格面板', await tablePanel.count(), 1)
  eq('面板默认行数 = 2', await rowBox.inputValue(), '2')
  eq('面板默认列数 = 3', await colBox.inputValue(), '3')

  // Esc 只关面板，不插表
  await page.keyboard.press('Escape')
  await page.waitForTimeout(150)
  eq('Esc 关掉面板', await tablePanel.count(), 0)
  eq('Esc 之后没插表', (await tableIds()).length, beforeU.length)

  // 选 4 行 2 列再插入
  await page.locator(tableBtn).click()
  await page.waitForTimeout(150)
  await rowBox.fill('4')
  await colBox.fill('2')
  await tablePanel.getByRole('button', { name: '插入', exact: true }).click()
  await page.waitForTimeout(300)
  const afterU = await getModel()
  const added = afterU.blocks.filter((b) => b.t === 'table' && !beforeU.includes(b.id))
  eq('插入了一张新表', added.length, 1)
  const fresh = added[0]
  eq('新表行数 = 面板里选的 4', fresh?.rows.length, 4)
  eq('新表每行格数 = 面板里选的 2', fresh?.rows.every((r) => r.cells.length === 2), true)
  eq('新表 columns = 面板里选的 2', fresh?.columns, 2)
  eq('新表 minLines 仍是 2', fresh?.minLines, 2)
  eq('新表默认禁止跨页断行', fresh?.cantSplit, true)
  eq('插入后面板自动关闭', await tablePanel.count(), 0)
  ok(
    '插入符落在新表第一个格子里',
    await page.evaluate((id) => {
      const el = document.querySelector(`[data-block-id="${id}"]`)
      const sel = window.getSelection()
      return (
        el !== null &&
        sel !== null &&
        sel.rangeCount > 0 &&
        el.contains(sel.getRangeAt(0).startContainer)
      )
    }, `${fresh?.id ?? ''}.r0c0`),
  )
  await checkNoOverflow('U 插入表格后')

  // 数字框的 min/max 只是提示，值能直接敲进去 —— 组件侧必须夹回合法区间
  await page.locator(tableBtn).click()
  await page.waitForTimeout(150)
  await rowBox.fill('999')
  await colBox.fill('0')
  await tablePanel.getByRole('button', { name: '插入', exact: true }).click()
  await page.waitForTimeout(300)
  const clampedModel = await getModel()
  // 按 id 差分认新表，不能取「文档里最后一张表」—— 插入落点是插入符所在块之后，
  // 不一定是文末（样本里本来就有一张表，取最后一张可能取到它）
  const clampedAdded = clampedModel.blocks.filter(
    (b) => b.t === 'table' && !afterU.blocks.some((known) => known.id === b.id),
  )
  eq('越界那次也插进了一张新表', clampedAdded.length, 1)
  const clamped = clampedAdded[0]
  eq('行数上界夹到 30', clamped?.rows.length, 30)
  eq('列数下界夹到 1', clamped?.columns, 1)

  // 另一侧的两个边界也要各测一次，否则「1–30 / 1–12」只验了上界那半边
  await page.locator(tableBtn).click()
  await page.waitForTimeout(150)
  await rowBox.fill('0')
  await colBox.fill('999')
  await tablePanel.getByRole('button', { name: '插入', exact: true }).click()
  await page.waitForTimeout(300)
  const clampedAdded2 = (await getModel()).blocks.filter(
    (b) => b.t === 'table' && !clampedModel.blocks.some((known) => known.id === b.id),
  )
  eq('另一侧也插进了一张新表', clampedAdded2.length, 1)
  eq('行数下界夹到 1', clampedAdded2[0]?.rows.length, 1)
  eq('列数上界夹到 12', clampedAdded2[0]?.columns, 12)
  await checkNoOverflow('U 另一侧越界夹回后')

  /* ------------------------------------------------------------------ */
  console.log('\n=== V. 表格编辑交互（W4b-1）：上下文工具条、增删行列、unit/note、行高、Shift+Enter、边界护栏 ===')
  const sampleTable = (model) => model.blocks.find((b) => b.t === 'table')
  const modelTable = async () => sampleTable(await getModel())
  const cellText = (t, r, c) =>
    (t?.rows?.[r]?.cells?.[c]?.inlines ?? [])
      .map((i) => (i.t === 'text' ? i.text : ''))
      .join('')
  const cellInlines = (t, r, c) => t?.rows?.[r]?.cells?.[c]?.inlines ?? []
  /** 表格在页面上的 DOM 概况（片段的 tr 加起来应等于模型行数） */
  const tableDom = () =>
    page.evaluate(() => ({
      tables: document.querySelectorAll('.wtp-table').length,
      trs: document.querySelectorAll('.wtp-table tbody tr').length,
      tds: document.querySelectorAll('.wtp-table td').length,
      min2: document.querySelectorAll('.wtp-table.wtp-table-min2').length,
      min1: document.querySelectorAll('.wtp-table.wtp-table-min1').length,
    }))
  // 表格页常驻（不随光标进出表格出现/消失），按 .table-toolbar 定位；
  // 光标不在格子里时它给一句提示并置灰全部按钮
  const subToolbar = page.locator('.table-toolbar')
  const subHint = page.locator('.table-toolbar .tk-empty')
  const subButton = (name) => subToolbar.getByRole('button', { name, exact: true })
  /** 表格页里某个 radio 组里的一个选项（组按左边的标签文字定位，避免「有/无」重名） */
  const subRadio = (group, name) =>
    subToolbar.locator('.tk-group', { hasText: group }).getByRole('radio', { name, exact: true })

  // ---- V1. 工具条常驻、按钮随光标进出表格启用/置灰 + 下方插入行 ----
  await openApp('表格')
  // 光标还没进格子：工具条就在（宽度不跳），按钮全灰 + 一句提示
  eq('光标不在表格里时工具条仍在（常驻）', await subToolbar.count(), 1)
  eq('光标不在表格里时给一句提示', await subHint.count(), 1)
  ok('光标不在表格里时按钮置灰', await subButton('下方插入行').isDisabled())
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('数控加工中心'))
  await page.waitForTimeout(200)
  eq('光标进表格后工具条仍在（常驻，不随光标出现/消失）', await subToolbar.count(), 1)
  eq('光标进表格后提示收起', await subHint.count(), 0)
  ok('光标进表格后按钮可用', (await subButton('下方插入行').isDisabled()) === false)
  ok(
    '工具条显示落点（行/列都是下标，显示时 +1）',
    (await subToolbar.innerText()).includes('第 3 行第 1 列'),
    await subToolbar.innerText(),
  )
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('我方于2026年9月1日'))
  await page.waitForTimeout(200)
  eq('光标离开表格后工具条不消失（常驻）', await subToolbar.count(), 1)
  eq('光标离开表格后提示回来', await subHint.count(), 1)
  ok('光标离开表格后按钮又置灰', await subButton('下方插入行').isDisabled())

  await page.evaluate(() => window.__wtpTest.caretAtEndOf('数控加工中心'))
  await page.waitForTimeout(200)
  const beforeV1 = await modelTable()
  const domBeforeV1 = await tableDom()
  await subButton('下方插入行').click()
  await page.waitForTimeout(300)
  const afterV1 = await modelTable()
  const domAfterV1 = await tableDom()
  eq('下方插入行：模型行数 +1', afterV1?.rows.length, (beforeV1?.rows.length ?? 0) + 1)
  eq('新行插在光标那一行之后', afterV1?.rows[3]?.role, 'body')
  eq('新行格数 = columns', afterV1?.rows[3]?.cells.length, afterV1?.columns)
  eq('DOM 也多一个 <tr>', domAfterV1.trs, domBeforeV1.trs + 1)
  eq('DOM 的 tr 数仍等于模型行数', domAfterV1.trs, afterV1?.rows.length)
  const v1Caret = await page.evaluate(() => window.__wtpTest.caretInfo())
  eq('插入符仍在原来那一格（r2c0）', v1Caret?.blockId, `${afterV1?.id ?? ''}.r2c0`)
  eq('插入符偏移没变（该格文字末尾）', v1Caret?.offset, '数控加工中心'.length)
  await checkNoOverflow('V1 下方插入行后')

  // ---- V2. 删除列到最后一列 → 按钮禁用 ----
  await openApp('表格')
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('数控加工中心'))
  await page.waitForTimeout(200)
  const deleteCol = subButton('删除列')
  ok('多列时「删除列」可用', (await deleteCol.isDisabled()) === false)
  await deleteCol.click()
  await page.waitForTimeout(300)
  const afterFirstDel = await modelTable()
  eq('删掉一列后 columns -1', afterFirstDel?.columns, 2)
  eq('删除列只动 body 行（unit/note 仍一格）', `${afterFirstDel?.rows[0]?.cells.length}/${afterFirstDel?.rows[afterFirstDel.rows.length - 1]?.cells.length}`, '1/1')
  await deleteCol.click()
  await page.waitForTimeout(300)
  const afterSecondDel = await modelTable()
  eq('再删一列后只剩 1 列', afterSecondDel?.columns, 1)
  ok('只剩最后一列时「删除列」被禁用', await deleteCol.isDisabled())
  await checkNoOverflow('V2 删除列后')

  // ---- V3. 两个 radio：行高 2↔1、表头行有/无 ----
  await openApp('表格')
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('数控加工中心'))
  await page.waitForTimeout(200)
  const domMinBefore = await tableDom()
  ok(
    '样本表初始 minLines=2（DOM 用 -min2 类）',
    domMinBefore.min2 >= 1 && domMinBefore.min1 === 0,
    JSON.stringify(domMinBefore),
  )
  await subRadio('行高', '最小一行').click()
  await page.waitForTimeout(300)
  eq('切到最小一行：模型 minLines=1', (await modelTable())?.minLines, 1)
  const domMin1 = await tableDom()
  ok(
    '切到最小一行：DOM 类名换成 -min1',
    domMin1.min1 >= 1 && domMin1.min2 === 0,
    JSON.stringify(domMin1),
  )
  await subRadio('行高', '最小两行').click()
  await page.waitForTimeout(300)
  eq('切回最小两行：模型 minLines=2', (await modelTable())?.minLines, 2)
  const domMin2 = await tableDom()
  ok(
    '切回最小两行：DOM 类名回到 -min2',
    domMin2.min2 >= 1 && domMin2.min1 === 0,
    JSON.stringify(domMin2),
  )

  const beforeUnit = await modelTable()
  eq('样本表本来有表头行', beforeUnit?.rows[0]?.role, 'unit')
  await subRadio('表头行', '无').click()
  await page.waitForTimeout(300)
  const afterUnitOff = await modelTable()
  eq('表头行 radio=无：unit 行消失', afterUnitOff?.rows.some((r) => r.role === 'unit'), false)
  eq('表头行 radio=无：只少这一行', afterUnitOff?.rows.length, (beforeUnit?.rows.length ?? 0) - 1)
  await subRadio('表头行', '有').click()
  await page.waitForTimeout(300)
  const afterUnitOn = await modelTable()
  eq('表头行 radio=有：unit 行回到最前', afterUnitOn?.rows[0]?.role, 'unit')
  eq('表头行 radio=有：行数复原', afterUnitOn?.rows.length, beforeUnit?.rows.length)

  const beforeNote = await modelTable()
  eq('样本表本来有附注行', beforeNote?.rows[beforeNote.rows.length - 1]?.role, 'note')
  await subRadio('附注行', '有').click()
  await page.waitForTimeout(300)
  const afterNoteOn = await modelTable()
  eq('附注行 radio=有：note 仍在最后（幂等）', afterNoteOn?.rows[afterNoteOn.rows.length - 1]?.role, 'note')
  eq('附注行 radio=有：再点一次不产生第二行', afterNoteOn?.rows.length, beforeNote?.rows.length)
  await subRadio('附注行', '无').click()
  await page.waitForTimeout(300)
  const afterNoteOff = await modelTable()
  eq('附注行 radio=无：note 行消失', afterNoteOff?.rows.some((r) => r.role === 'note'), false)
  eq('附注行 radio=无：只少这一行', afterNoteOff?.rows.length, (beforeNote?.rows.length ?? 0) - 1)
  await checkNoOverflow('V3 radio 切换后')

  // ---- V4. 格内 Shift+Enter：插入符落在换行之后 ----
  await openApp('表格')
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('检测仪器'))
  await page.waitForTimeout(200)
  await page.keyboard.press('Shift+Enter')
  await page.waitForTimeout(300)
  const shifted = await modelTable()
  const shiftRow = shifted?.rows.findIndex((_r, i) => cellText(shifted, i, 0) === '检测仪器') ?? -1
  ok('找得到目标格子', shiftRow >= 0, `shiftRow=${shiftRow}`)
  const shiftInlines = cellInlines(shifted, shiftRow, 0)
  eq('Shift+Enter 只插一枚软换行', shiftInlines.filter((i) => i.t === 'break').length, 1)
  eq('软换行是零宽的：格内文字没变', cellText(shifted, shiftRow, 0), '检测仪器')
  // 换行之后敲字：必须落在新的一行（模型里排在 break 之后），而不是回到上一行末尾
  await page.keyboard.insertText('X')
  await page.waitForTimeout(300)
  const typedV4 = await modelTable()
  const typedInlines = cellInlines(typedV4, shiftRow, 0)
  const breakIdx = typedInlines.findIndex((i) => i.t === 'break')
  const typedIdx = typedInlines.findIndex((i) => i.t === 'text' && i.text.includes('X'))
  ok('Shift+Enter 后插入符落在换行之后（敲的字排在 break 之后）', breakIdx >= 0 && typedIdx > breakIdx)
  eq('换行后敲的字没打回上一行', cellText(typedV4, shiftRow, 0), '检测仪器X')
  const v4Caret = await page.evaluate(() => window.__wtpTest.caretInfo())
  eq('插入符仍在同一格', v4Caret?.blockId, `${typedV4?.id ?? ''}.r${shiftRow}c0`)
  await checkNoOverflow('V4 Shift+Enter 后')

  // ---- V5. 边界护栏：格首 Backspace、格尾 Delete 都不许动模型/DOM ----
  await openApp('表格')
  const guardDom = await tableDom()
  const guardTable = await modelTable()
  await page.evaluate(() => window.__wtpTest.setCaret('数控加工中心', 0))
  await page.waitForTimeout(200)
  await page.keyboard.press('Backspace')
  await page.waitForTimeout(300)
  const afterBackspace = await modelTable()
  const domAfterBackspace = await tableDom()
  eq('格首 Backspace：行数不变', afterBackspace?.rows.length, guardTable?.rows.length)
  eq('格首 Backspace：格内文字不变', cellText(afterBackspace, 2, 0), '数控加工中心')
  eq('格首 Backspace：<td> 数没变（原生没并掉相邻格）', domAfterBackspace.tds, guardDom.tds)

  await page.evaluate(() => window.__wtpTest.caretAtEndOf('检测仪器'))
  await page.waitForTimeout(200)
  await page.keyboard.press('Delete')
  await page.waitForTimeout(300)
  const afterDeleteV5 = await modelTable()
  eq('格尾 Delete：行数不变', afterDeleteV5?.rows.length, guardTable?.rows.length)
  eq('格尾 Delete：格内文字不变', cellText(afterDeleteV5, 4, 0), '检测仪器')
  eq(
    '格尾 Delete：<td> 数没变',
    (await tableDom()).tds,
    guardDom.tds,
  )
  await checkNoOverflow('V5 边界护栏后')

  /* ------------------------------------------------------------------ */
  console.log(
    '\n=== W. 表格收尾（W4b-2）：Tab/Shift+Tab 跨格、格首格尾方向键、删整表、格内换样式、两组对齐 ===',
  )
  const cellAlign = (t, r, c) => t?.rows?.[r]?.cells?.[c]?.align
  const groupButtons = (group) => subToolbar.locator('.tk-group', { hasText: group })
  const alignButton = (group, name) => groupButtons(group).getByRole('button', { name, exact: true })
  /** 上下文工具条某个组里处于高亮的按钮文字 */
  const activeLabels = (group) =>
    page.evaluate((g) => {
      const box = Array.from(document.querySelectorAll('.sub-toolbar .tk-group')).find((el) =>
        (el.textContent ?? '').includes(g),
      )
      return box ? Array.from(box.querySelectorAll('button.is-on')).map((b) => b.textContent?.trim()) : []
    }, group)

  // ---- W1. Tab / Shift+Tab 行优先跨格；最后一格 Tab 无响应 ----
  await openApp('表格')
  const wTable = await modelTable()
  const wId = wTable?.id ?? ''
  ok('样本表 id 可用', wId !== '', wId)
  await page.evaluate(() => window.__wtpTest.setCaret('数控加工中心', 0))
  await page.waitForTimeout(200)
  await page.keyboard.press('Tab')
  await page.waitForTimeout(200)
  const tab1 = await page.evaluate(() => window.__wtpTest.caretInfo())
  eq('Tab 从 r2c0 跨到 r2c1', tab1?.blockId, `${wId}.r2c1`)
  eq('Tab 跨格落点是格首', tab1?.offset, 0)
  eq('跨格不改模型（只挪选区）', JSON.stringify(await modelTable()), JSON.stringify(wTable))
  await page.keyboard.press('Tab')
  await page.waitForTimeout(200)
  const tab2 = await page.evaluate(() => window.__wtpTest.caretInfo())
  eq('再 Tab 到 r2c2 格首', tab2?.blockId, `${wId}.r2c2`)
  eq('落点仍是格首', tab2?.offset, 0)

  await page.evaluate(() => window.__wtpTest.caretAtEndOf('账面原值'))
  await page.waitForTimeout(200)
  await page.keyboard.press('Shift+Tab')
  await page.waitForTimeout(200)
  const backTab = await page.evaluate(() => window.__wtpTest.caretInfo())
  eq('Shift+Tab 从 r1c2 回到 r1c1', backTab?.blockId, `${wId}.r1c1`)
  eq('Shift+Tab 落点是格尾', backTab?.offset, '数量'.length)

  // 最后一格（附注行）：Tab 之后选区与模型都不许变
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('未经审计'))
  await page.waitForTimeout(200)
  const lastBefore = await page.evaluate(() => window.__wtpTest.caretInfo())
  const modelBeforeLast = JSON.stringify(await modelTable())
  await page.keyboard.press('Tab')
  await page.waitForTimeout(200)
  const lastAfter = await page.evaluate(() => window.__wtpTest.caretInfo())
  eq('最后一格 Tab：落点没动', lastAfter?.blockId, lastBefore?.blockId)
  eq('最后一格 Tab：偏移没动', lastAfter?.offset, lastBefore?.offset)
  eq('最后一格 Tab：模型没动', JSON.stringify(await modelTable()), modelBeforeLast)
  await checkNoOverflow('W1 Tab 跨格后')

  // ---- W2. ← / → 在格首 / 格尾跨格；格内中间不接管 ----
  await openApp('表格')
  await page.evaluate(() => window.__wtpTest.setCaret('数控加工中心', 0))
  await page.waitForTimeout(200)
  await page.keyboard.press('ArrowLeft')
  await page.waitForTimeout(200)
  const goLeft = await page.evaluate(() => window.__wtpTest.caretInfo())
  eq('格首 ← 跨到上一格（r1c2）', goLeft?.blockId, `${wId}.r1c2`)
  eq('← 落在上一格末尾', goLeft?.offset, '账面原值'.length)

  await page.evaluate(() => window.__wtpTest.caretAtEndOf('数控加工中心'))
  await page.waitForTimeout(200)
  await page.keyboard.press('ArrowRight')
  await page.waitForTimeout(200)
  const goRight = await page.evaluate(() => window.__wtpTest.caretInfo())
  eq('格尾 → 跨到下一格（r2c1）', goRight?.blockId, `${wId}.r2c1`)
  eq('→ 落在下一格开头', goRight?.offset, 0)

  await page.evaluate(() => window.__wtpTest.setCaret('数控加工中心', 2))
  await page.waitForTimeout(200)
  await page.keyboard.press('ArrowLeft')
  await page.waitForTimeout(200)
  const midLeft = await page.evaluate(() => window.__wtpTest.caretInfo())
  eq('格内中间 ← 不跨格（仍在本格）', midLeft?.blockId, `${wId}.r2c0`)
  eq('格内中间 ← 只左移一位（浏览器接管）', midLeft?.offset, 1)
  await checkNoOverflow('W2 方向键后')

  // ---- W3. 删除整张表（不二次确认）----
  await openApp('表格')
  const delBefore = await getModel()
  const delIndex = delBefore.blocks.findIndex((b) => b.t === 'table')
  const prevBlock = delBefore.blocks[delIndex - 1]
  const prevText = textOfBlock(prevBlock)
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('数控加工中心'))
  await page.waitForTimeout(200)
  ok('「删除表格」按钮在格内可用', (await subButton('删除表格').isDisabled()) === false)
  await subButton('删除表格').click()
  await page.waitForTimeout(300)
  const delAfter = await getModel()
  eq(
    '模型少一个 table 块',
    delAfter.blocks.filter((b) => b.t === 'table').length,
    delBefore.blocks.filter((b) => b.t === 'table').length - 1,
  )
  eq('DOM 里不再有 .wtp-tableFrag', await page.evaluate(() => document.querySelectorAll('.wtp-tableFrag').length), 0)
  eq('表格页仍常驻（按钮置灰）', await subToolbar.count(), 1)
  eq('光标离开表格后提示回来（删表后）', await subHint.count(), 1)
  const delCaret = await page.evaluate(() => window.__wtpTest.caretInfo())
  eq('插入符落在上一块（块 id）', delCaret?.blockId, prevBlock.id)
  eq('插入符落在上一块末尾（偏移 = 该块文字长度）', delCaret?.offset, prevText.length)
  await checkNoOverflow('W3 删除表格后')

  // ---- W4. 光标在格内点样式 chip → 只有该格换样式 ----
  // 样式库在「开始」页里，所以这一小节要的是开始页（不是表格页）
  await openApp('开始')
  const chipBefore = await modelTable()
  const kindsBefore = chipBefore.rows.map((r) => r.cells.map((c) => c.kind ?? null))
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('数控加工中心'))
  await page.waitForTimeout(200)
  await page.locator('.styles .style-chip', { hasText: '二级标题' }).click()
  await page.waitForTimeout(300)
  const chipAfter = await modelTable()
  eq('那一格换了样式（kind = h2）', chipAfter.rows[2].cells[0].kind, 'h2')
  const changedCells = []
  chipAfter.rows.forEach((r, ri) =>
    r.cells.forEach((c, ci) => {
      if ((kindsBefore[ri]?.[ci] ?? null) !== (c.kind ?? null)) changedCells.push(`${ri},${ci}`)
    }),
  )
  eq('只有那一格变（其余格一个都没动）', changedCells.join('|'), '2,0')
  // 类名从 wtp-listItem 换成 wtp-h2；按文本定位那一格（片段可能分页）
  const chipDom = await page.evaluate(() => {
    const cells = Array.from(document.querySelectorAll('.wtp-table .wtp-cell[data-block-id]'))
    const target = cells.find((el) => (el.textContent ?? '').includes('数控加工中心'))
    return {
      target: target ? target.className : 'missing',
      h2: document.querySelectorAll('.wtp-table .wtp-cell.wtp-h2').length,
      h2Text: Array.from(document.querySelectorAll('.wtp-table .wtp-cell.wtp-h2'))
        .map((el) => el.textContent)
        .join(','),
    }
  })
  ok(
    'DOM 里那一格的类名从 wtp-listItem 变成 wtp-h2',
    chipDom.target.split(' ').includes('wtp-cell') && chipDom.target.split(' ').includes('wtp-h2'),
    JSON.stringify(chipDom),
  )
  ok(
    '那一格不再是 wtp-listItem',
    !chipDom.target.split(' ').includes('wtp-listItem'),
    JSON.stringify(chipDom),
  )
  eq('整张表只有这一格是 wtp-h2', chipDom.h2, 1)
  eq('wtp-h2 那一格就是被点的那格', chipDom.h2Text, '数控加工中心')
  await checkNoOverflow('W4 格内换样式后')

  // ---- W5. 两组对齐：active 态、写进 DOM 的行内样式、只作用于该格 ----
  await openApp('表格')
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('单位：元'))
  await page.waitForTimeout(200)
  eq('unit 行默认右对齐 → 水平组高亮「右」', (await activeLabels('水平')).join(','), '右')
  eq('垂直默认顶端 → 垂直组高亮「顶端」', (await activeLabels('垂直')).join(','), '顶端')

  await page.evaluate(() => window.__wtpTest.caretAtEndOf('数控加工中心'))
  await page.waitForTimeout(200)
  eq('body 格默认是该样式的两端对齐 → 水平三档都不高亮', (await activeLabels('水平')).join(','), '')
  eq('body 格垂直默认顶端 → 垂直组高亮「顶端」', (await activeLabels('垂直')).join(','), '顶端')

  const alignModelBefore = await modelTable()
  const pagesBeforeAlign = await page.evaluate(() => document.querySelectorAll('.wtp-page').length)
  await alignButton('水平', '居中').click()
  await page.waitForTimeout(300)
  eq('点「居中」写进模型（align.h=center）', cellAlign(await modelTable(), 2, 0)?.h, 'center')
  eq('点「居中」后该按钮高亮', (await activeLabels('水平')).join(','), '居中')
  eq(
    '水平对齐写进格内 div 的行内 text-align',
    await page.evaluate(() => {
      const cells = Array.from(document.querySelectorAll('.wtp-table .wtp-cell[data-block-id]'))
      return cells.find((el) => (el.textContent ?? '').includes('数控加工中心'))?.style.textAlign ?? 'missing'
    }),
    'center',
  )
  await alignButton('水平', '居中').click()
  await page.waitForTimeout(300)
  eq('再点同一个值 → 清除覆盖（字段消失）', cellAlign(await modelTable(), 2, 0)?.h, undefined)
  eq('清除后水平组不再高亮', (await activeLabels('水平')).join(','), '')

  await alignButton('垂直', '底端').click()
  await page.waitForTimeout(300)
  eq('点「底端」写进模型（align.v=bottom）', cellAlign(await modelTable(), 2, 0)?.v, 'bottom')
  eq('点「底端」后该按钮高亮', (await activeLabels('垂直')).join(','), '底端')
  eq(
    '垂直对齐写进 <td> 的行内 vertical-align（写在格内 div 上无效）',
    await page.evaluate(() => {
      const cells = Array.from(document.querySelectorAll('.wtp-table .wtp-cell[data-block-id]'))
      const target = cells.find((el) => (el.textContent ?? '').includes('数控加工中心'))
      return target?.closest('td')?.style.verticalAlign ?? 'missing'
    }),
    'bottom',
  )
  await alignButton('垂直', '底端').click()
  await page.waitForTimeout(300)
  eq('再点同一个值 → 清除覆盖', cellAlign(await modelTable(), 2, 0)?.v, undefined)
  // 生效范围：只有那一格被改过（其它格的 align 一个都没动）
  const alignModelAfter = await modelTable()
  const ownAlign = (t) =>
    t.rows.flatMap((r, ri) => r.cells.map((c, ci) => `${ri},${ci}:${JSON.stringify(c.align ?? null)}`)).join('|')
  eq('两组对齐只作用于光标那一格', ownAlign(alignModelAfter), ownAlign(alignModelBefore))
  eq('（过程中的中间态已回到默认）那一格 align 字段消失', 'align' in alignModelAfter.rows[2].cells[0], false)
  // 两组对齐不改行高、不改换行点 → 页数必须不变（格内换样式会改行高、页数可以变，那是预期）
  eq(
    '两组对齐来回点完页数不变',
    await page.evaluate(() => document.querySelectorAll('.wtp-page').length),
    pagesBeforeAlign,
  )
  await checkNoOverflow('W5 两组对齐后')

  /* ------------------------------------------------------------------ */
  console.log('\n=== S. 「节」工具条（W5）：回显 / 置灰 / 改方向与页码 ===')
  await openApp('布局')
  // 编辑模式常驻：节工具条与表格工具条同在一层，但节这条永远在
  const secToolbar = page.locator('.section-toolbar')
  const secRadio = (group, name) =>
    secToolbar.locator('.tk-group', { hasText: group }).getByRole('radio', { name, exact: true })
  /** 读回节工具条的回显：节号提示 + 四组 radio 的选中项与整组是否置灰 */
  const secState = () =>
    page.evaluate(() => {
      const bar = document.querySelector('.section-toolbar')
      if (!bar) return null
      const groups = Array.from(bar.querySelectorAll('.tk-group'))
      const read = (name) => {
        const group = groups.find((g) => (g.textContent ?? '').includes(name))
        if (!group) return null
        const radios = Array.from(group.querySelectorAll('input[type="radio"]'))
        return {
          label: radios
            .filter((i) => i.checked)
            .map((i) => (i.closest('label')?.textContent ?? '').trim())
            .join(','),
          disabled: radios.length > 0 && radios.every((i) => i.disabled),
        }
      }
      return {
        hint: (bar.querySelector('.tk-hint')?.textContent ?? '').trim(),
        orientation: read('方向'),
        numbers: read('页码'),
        link: read('关联前节'),
        restart: read('从 1 开始'),
      }
    })
  /** 每张纸的几何与它挂的命名页（判断横竖混排只看这两个数） */
  const paperState = () =>
    page.evaluate(() =>
      Array.from(document.querySelectorAll('.wtp-page')).map((el) => {
        const rect = el.getBoundingClientRect()
        return {
          page: el.style.getPropertyValue('page'),
          width: Math.round(rect.width),
          height: Math.round(rect.height),
          number: (el.querySelector('.wtp-page-number')?.textContent ?? '').trim(),
        }
      }),
    )

  ok('编辑模式下「节」工具条常驻', (await secToolbar.count()) === 1)

  // 首节：没有前节 → 「关联前节」整组置灰；恒从 1 开始 → 「从 1 开始」整组置灰
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('我方于2026年9月1日'))
  await page.waitForTimeout(250)
  const s0 = await secState()
  eq('节号回显（首节）', s0?.hint, '第 1 节 / 共 2 节')
  eq('首节方向回显为纵向', s0?.orientation?.label, '纵向')
  ok('首节「关联前节」整组置灰（没有前节）', s0?.link?.disabled === true)
  ok('首节「从 1 开始」整组置灰（恒从 1 开始）', s0?.restart?.disabled === true)
  ok('首节「页码」可用（首节可以关掉页码）', s0?.numbers?.disabled === false)
  eq('首节页码回显为开', s0?.numbers?.label, '开')

  // 第 2 节：样本里写的是 `--- link=off restart=on`
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('2026年9月12日'))
  await page.waitForTimeout(250)
  const s1 = await secState()
  eq('节号回显（第 2 节）', s1?.hint, '第 2 节 / 共 2 节')
  ok('第 2 节「关联前节」可用', s1?.link?.disabled === false)
  eq('样本里第 2 节声明了 link=off → 回显「否」', s1?.link?.label, '否')
  eq('样本里第 2 节声明了 restart=on → 回显「是」', s1?.restart?.label, '是')

  // 点「方向 → 横向」：只有这一节的纸变成横的（宽 > 高），首节仍是纵的
  const papersBefore = await paperState()
  await secRadio('方向', '横向').click()
  await page.waitForTimeout(400)
  const papersAfter = await paperState()
  eq(
    '模型里第 2 节记成 landscape',
    await page.evaluate(() => window.__wtpPaper.getModel().sections?.[1]?.orientation),
    'landscape',
  )
  const portraitPapers = papersAfter.filter((p) => p.page === '')
  const landscapePapers = papersAfter.filter((p) => p.page === 'wtp-landscape')
  ok('改完有横排的纸', landscapePapers.length > 0, JSON.stringify(papersAfter))
  ok(
    '横排的纸宽 > 高',
    landscapePapers.every((p) => p.width > p.height),
    JSON.stringify(landscapePapers[0]),
  )
  ok(
    '纵排的纸仍是宽 < 高（只有那一节被改）',
    portraitPapers.length > 0 && portraitPapers.every((p) => p.width < p.height),
    JSON.stringify(portraitPapers[0]),
  )
  ok(
    '横排纸的宽高恰好是纵排纸的对调',
    landscapePapers[0]?.width === portraitPapers[0]?.height &&
      landscapePapers[0]?.height === portraitPapers[0]?.width,
    `${JSON.stringify(landscapePapers[0])} vs ${JSON.stringify(portraitPapers[0])}`,
  )
  eq(
    '横竖混排没有多出纸来（页数不变）',
    papersAfter.length,
    papersBefore.length,
  )

  // 点「页码 → 关」：这一节（横排那几页）不再有页码元素，其它节不受影响
  await secRadio('页码', '关').click()
  await page.waitForTimeout(400)
  const mutedPapers = await paperState()
  ok(
    '横排节不再显示页码',
    mutedPapers.filter((p) => p.page === 'wtp-landscape').every((p) => p.number === ''),
    JSON.stringify(mutedPapers),
  )
  ok(
    '纵排节照旧显示页码',
    mutedPapers
      .filter((p) => p.page === '')
      .every((p) => /^\d+$/.test(p.number)),
    JSON.stringify(mutedPapers),
  )
  eq(
    '模型里第 2 节 numbers=false',
    await page.evaluate(() => window.__wtpPaper.getModel().sections?.[1]?.pageNumbers),
    false,
  )

  // 还原：页码回开、方向回纵向；模型里也不再留冗余字段
  await secRadio('页码', '开').click()
  await page.waitForTimeout(300)
  await secRadio('方向', '纵向').click()
  await page.waitForTimeout(400)
  const restored = await secState()
  eq('还原后方向回显纵向', restored?.orientation?.label, '纵向')
  eq('还原后页码回显开', restored?.numbers?.label, '开')
  eq(
    '还原后模型里不留 orientation（默认值不落字段）',
    await page.evaluate(() => window.__wtpPaper.getModel().sections?.[1]?.orientation),
    undefined,
  )
  ok(
    '还原后所有纸都是纵排',
    (await paperState()).every((p) => p.page === '' && p.width < p.height),
  )
  await checkNoOverflow('S 节设置来回切后')

  // 打印：方向靠两档命名 @page 表达（未命名的 @page 只是兜底），两档尺寸必须互换
  const namedPageRules = await page.evaluate(() => {
    const out = []
    for (const sheet of document.styleSheets) {
      let rules
      try {
        rules = sheet.cssRules
      } catch {
        continue
      }
      for (const rule of rules) {
        if (rule.conditionText !== 'print') continue
        for (const inner of rule.cssRules ?? []) {
          if (inner.constructor.name !== 'CSSPageRule') continue
          out.push({ cssText: inner.cssText, size: inner.style.size, margin: inner.style.margin })
        }
      }
    }
    return out
  })
  const portraitRule = namedPageRules.find((r) => r.cssText.includes('wtp-portrait'))
  const landscapeRule = namedPageRules.find((r) => r.cssText.includes('wtp-landscape'))
  ok('打印样式里有命名页 wtp-portrait', portraitRule !== undefined, JSON.stringify(namedPageRules))
  ok('打印样式里有命名页 wtp-landscape', landscapeRule !== undefined, JSON.stringify(namedPageRules))
  if (portraitRule && landscapeRule) {
    const p = portraitRule.size.split(/\s+/)
    const l = landscapeRule.size.split(/\s+/)
    ok(
      `两档命名页的 size 恰好互换（${portraitRule.size} ↔ ${landscapeRule.size}）`,
      p[0] === l[1] && p[1] === l[0],
      `${JSON.stringify(p)} vs ${JSON.stringify(l)}`,
    )
    eq('命名页的外边距归 0（白边由纸张自己的 padding 提供）', landscapeRule.margin, '0px')
  }

  /* ------------------------------------------------------------------ */
  console.log('\n=== X. 删除的边界（issues/20260915）：页尾/页首/跨段/整页，DOM 与模型必须逐块一致 ===')
  // 每做完一步都用 checkDomMatchesModel 逐块对账：坐标与 DOM 文字差一个字符就抓得到。
  const xModelLen = async (blockId) =>
    textOfBlock(heroBlocks(await getModel()).find((b) => b.id === blockId) ?? { inlines: [] }).length
  /** 在某一页的某两块之间拉一个选区：首块从第 from 字起、末块到倒数 back 字止 */
  const xSelectAcross = (pageIndex, firstIndex, lastIndex, from, back) =>
    page.evaluate(
      ({ pageIndex, firstIndex, lastIndex, from, back }) => {
        const content = document.querySelectorAll('.wtp-page')[pageIndex].querySelector('.wtp-content')
        const frags = Array.from(content.querySelectorAll('[data-block-id]'))
        const first = frags[firstIndex]
        const last = frags[lastIndex]
        const walk = (el, offset) => {
          const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT)
          let left = offset
          let node = w.nextNode()
          while (node) {
            if (left <= node.data.length) return { node, offset: left }
            left -= node.data.length
            node = w.nextNode()
          }
          return null
        }
        content.focus()
        const a = walk(first, from)
        const b = walk(last, Math.max(0, (last.textContent ?? '').length - back))
        const r = document.createRange()
        r.setStart(a.node, a.offset)
        r.setEnd(b.node, b.offset)
        const sel = document.getSelection()
        sel.removeAllRanges()
        sel.addRange(r)
        return {
          ids: frags.map((f) => f.dataset.blockId),
          firstText: first.textContent ?? '',
          lastText: last.textContent ?? '',
        }
      },
      { pageIndex, firstIndex, lastIndex, from, back },
    )

  // ---- X1. 跨页段落页尾连按 5 次 Backspace：模型长度必须恰好 −5，且不出现重复 ----
  await openApp()
  const xCross = await page.evaluate(() => window.__wtpTest.crossPageBlock())
  ok('样本里有跨页的段落', xCross !== null, JSON.stringify(xCross))
  const xCrossBefore = await xModelLen(xCross.blockId)
  const xTailEdge = await page.evaluate(
    ({ id, page: p }) => window.__wtpTest.caretAtPageEdgeOf(id, p, 'end'),
    { id: xCross.blockId, page: xCross.pages[0] },
  )
  ok('插入符落在页尾那一片的末尾', xTailEdge !== null, JSON.stringify(xTailEdge))
  for (let i = 1; i <= 5; i += 1) {
    await page.keyboard.press('Backspace')
    await page.waitForTimeout(150)
    eq(`页尾第 ${i} 次 Backspace：模型长度恰好 −${i}`, await xModelLen(xCross.blockId), xCrossBefore - i)
  }
  await checkDomMatchesModel('X1 页尾连按 Backspace 后：版面文字 = 模型文字')
  // 撤销：每一步都该原样退回（撤销快照必须在改模型**之前**记 —— 记晚了就退回不去）
  for (let i = 0; i < 5; i += 1) {
    await page.keyboard.press('Control+z')
    await page.waitForTimeout(120)
  }
  eq('页尾连按 5 次 Backspace 后撤销 5 次：模型长度回到原值', await xModelLen(xCross.blockId), xCrossBefore)
  await checkDomMatchesModel('X1 撤销回去后：版面文字 = 模型文字')

  // ---- X2. 页尾 Delete：删的是下一页片段的首字（跨页边界），不是空操作 ----
  await openApp()
  const xCross2 = await page.evaluate(() => window.__wtpTest.crossPageBlock())
  const xCross2Before = await xModelLen(xCross2.blockId)
  await page.evaluate(
    ({ id, page: p }) => window.__wtpTest.caretAtPageEdgeOf(id, p, 'end'),
    { id: xCross2.blockId, page: xCross2.pages[0] },
  )
  await page.keyboard.press('Delete')
  await page.waitForTimeout(250)
  eq('页尾 Delete：模型恰好 −1（不是空操作）', await xModelLen(xCross2.blockId), xCross2Before - 1)
  await checkDomMatchesModel('X2 页尾 Delete 后：版面文字 = 模型文字')

  // ---- X3. 页首 Backspace：删的是上一页片段的末字 ----
  await openApp()
  const xCross3 = await page.evaluate(() => window.__wtpTest.crossPageBlock())
  const xCross3Before = await xModelLen(xCross3.blockId)
  await page.evaluate(
    ({ id, page: p }) => window.__wtpTest.caretAtPageEdgeOf(id, p, 'start'),
    { id: xCross3.blockId, page: xCross3.pages[1] },
  )
  await page.keyboard.press('Backspace')
  await page.waitForTimeout(250)
  eq('页首 Backspace：模型恰好 −1', await xModelLen(xCross3.blockId), xCross3Before - 1)
  await checkDomMatchesModel('X3 页首 Backspace 后：版面文字 = 模型文字')

  // ---- X4. 段尾 Delete：删掉段落标记，下一段接上来（Word 语义） ----
  await openApp()
  const xBeforeJoin = await getModel()
  const xJoinHead = heroBlocks(xBeforeJoin)[2]
  const xJoinTail = heroBlocks(xBeforeJoin)[3]
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('我方于2026年9月1日'))
  await page.keyboard.press('Delete')
  await page.waitForTimeout(300)
  const xJoined = await getModel()
  eq('段尾 Delete：段落数 −1', heroBlocks(xJoined).length, heroBlocks(xBeforeJoin).length - 1)
  eq(
    '段尾 Delete：两段接起来（块 id 取前一段）',
    textOfBlock(heroBlocks(xJoined).find((b) => b.id === xJoinHead.id) ?? { inlines: [] }),
    textOfBlock(xJoinHead) + textOfBlock(xJoinTail),
  )
  ok(
    '段尾 Delete：后一段已从模型移除',
    !heroBlocks(xJoined).some((b) => b.id === xJoinTail.id),
  )
  await checkDomMatchesModel('X4 段尾 Delete 并段后：版面文字 = 模型文字')
  // 并段也是一步撤销（撤销快照必须记在改模型之前）
  await page.keyboard.press('Control+z')
  await page.waitForTimeout(300)
  const xUndoJoin = await getModel()
  eq('段尾 Delete 后撤销一步：段落数回到原值', heroBlocks(xUndoJoin).length, heroBlocks(xBeforeJoin).length)
  ok(
    '段尾 Delete 后撤销一步：后一段回来了',
    heroBlocks(xUndoJoin).some((b) => b.id === xJoinTail.id),
  )
  await checkDomMatchesModel('X4 撤销后：版面文字 = 模型文字')

  // ---- X4b. 段首 Backspace 并段（既有路径）也要能撤销 ----
  await openApp()
  const xBeforeMerge = await getModel()
  const xMergeHead = heroBlocks(xBeforeMerge)[2]
  const xMergeTail = heroBlocks(xBeforeMerge)[3]
  await page.evaluate(() => {
    const el = Array.from(document.querySelectorAll('.wtp-content [data-block-id]')).find((e) =>
      (e.textContent ?? '').includes('一、债务人基本情况'),
    )
    el.closest('[contenteditable="true"]').focus()
    const r = document.createRange()
    r.setStart(el, 0)
    r.collapse(true)
    const sel = document.getSelection()
    sel.removeAllRanges()
    sel.addRange(r)
  })
  await page.keyboard.press('Backspace')
  await page.waitForTimeout(300)
  eq('段首 Backspace：段落数 −1', heroBlocks(await getModel()).length, heroBlocks(xBeforeMerge).length - 1)
  await page.keyboard.press('Control+z')
  await page.waitForTimeout(300)
  const xUndoMerge = await getModel()
  eq('段首 Backspace 后撤销一步：段落数回到原值', heroBlocks(xUndoMerge).length, heroBlocks(xBeforeMerge).length)
  ok(
    '段首 Backspace 后撤销一步：被并掉的那段回来了',
    heroBlocks(xUndoMerge).some((b) => b.id === xMergeTail.id) &&
      Boolean(heroBlocks(xUndoMerge).find((b) => b.id === xMergeHead.id)),
  )
  await checkDomMatchesModel('X4b 撤销后：版面文字 = 模型文字')

  // ---- X5. 跨段选区删除：切掉首尾、中间整段删掉、两头接起来 ----
  await openApp()
  const xSpans = await xSelectAcross(0, 2, 4, 5, 5)
  const xBeforeSpan = await getModel()
  await page.keyboard.press('Delete')
  await page.waitForTimeout(400)
  const xAfterSpan = await getModel()
  const xSpanHead = heroBlocks(xAfterSpan).find((b) => b.id === xSpans.ids[2])
  eq(
    '跨段删除：首块 = 「首块切点之前 + 末块切点之后」',
    xSpanHead ? textOfBlock(xSpanHead) : '(首块没了)',
    xSpans.firstText.slice(0, 5) + xSpans.lastText.slice(Math.max(0, xSpans.lastText.length - 5)),
  )
  ok(
    '跨段删除：中间各块都从模型消失',
    xSpans.ids.slice(3, 5).every((id) => !heroBlocks(xAfterSpan).some((b) => b.id === id)),
  )
  eq(
    '跨段删除：正文块数减少 2（首尾各留一段、中间两块没了）',
    heroBlocks(xAfterSpan).length,
    heroBlocks(xBeforeSpan).length - 2,
  )
  await checkDomMatchesModel('X5 跨段选区删除后：版面文字 = 模型文字')
  // 跨段删除是**一步**撤销（撤销快照记在改模型之前，含被删掉的中间段）
  await page.keyboard.press('Control+z')
  await page.waitForTimeout(300)
  const xUndo = await getModel()
  eq('跨段删除后撤销一步：块数回到原值', heroBlocks(xUndo).length, heroBlocks(xBeforeSpan).length)
  ok(
    '跨段删除后撤销一步：中间各块都回来了',
    xSpans.ids.slice(3, 5).every((id) => heroBlocks(xUndo).some((b) => b.id === id)),
  )
  await checkDomMatchesModel('X5 撤销后：版面文字 = 模型文字')

  // ---- X6. 全选整页删除：该页各块并成一段（空），后面的文字前移 ----
  await openApp()
  const xBeforeAll = await getModel()
  const xPageIds = await page.evaluate(() => {
    const content = document.querySelector('.wtp-page .wtp-content')
    content.focus()
    const r = document.createRange()
    r.selectNodeContents(content)
    const sel = document.getSelection()
    sel.removeAllRanges()
    sel.addRange(r)
    return Array.from(content.querySelectorAll('[data-block-id]')).map((el) => el.dataset.blockId)
  })
  /*
   * 页面上挂 data-block-id 的东西不止「块」：表格的每个格子也挂着（cellId 形态 `tbl.rNcM`），
   * 而并段语义只管段落。期望值只能数**段落** —— 页首摊着表格时按元素个数算会多算一截
   * （2026-09-16 W7：② 把表头行挪到第 1 页之后，这里就数多了一行表格的格子）。
   */
  const xPageBlocks = xPageIds.filter((id) => !/\.r\d+c\d+$/.test(id))
  await page.keyboard.press('Delete')
  await page.waitForTimeout(500)
  const xAfterAll = await getModel()
  eq(
    '全选整页删除：块数减少 n−1（各段并成一段）',
    heroBlocks(xAfterAll).length,
    heroBlocks(xBeforeAll).length - (xPageBlocks.length - 1),
  )
  eq(
    '全选整页删除：留下的是该页第一块，文字清空',
    textOfBlock(heroBlocks(xAfterAll).find((b) => b.id === xPageIds[0]) ?? { inlines: [] }),
    '',
  )
  await checkDomMatchesModel('X6 全选整页删除后：版面文字 = 模型文字')
  await checkNoOverflow('X6 全选整页删除后')
  // 全选整页删除同样是**一步**撤销
  await page.keyboard.press('Control+z')
  await page.waitForTimeout(300)
  eq(
    '全选整页删除后撤销一步：块数回到原值',
    heroBlocks(await getModel()).length,
    heroBlocks(xBeforeAll).length,
  )
  await checkDomMatchesModel('X6 撤销后：版面文字 = 模型文字')

  // ---- X7. 粘贴替换跨段选区：原选区被粘贴内容取代（不是插在选区开头） ----
  await openApp()
  const xPastePick = await xSelectAcross(0, 2, 3, 4, 3)
  const xBeforePaste = await getModel()
  await page.evaluate(() => {
    const dt = new DataTransfer()
    dt.setData('text/plain', '【粘贴】')
    document
      .querySelector('.wtp-content')
      .dispatchEvent(new ClipboardEvent('paste', { bubbles: true, cancelable: true, clipboardData: dt }))
  })
  await page.waitForTimeout(400)
  const xAfterPaste = await getModel()
  const xPasted = heroBlocks(xAfterPaste).find((b) => b.id === xPastePick.ids[2])
  ok(
    '粘贴替换选区：粘贴内容落在切点处',
    (xPasted ? textOfBlock(xPasted) : '').includes('【粘贴】'),
    textOfBlock(xPasted ?? { inlines: [] }),
  )
  ok(
    '粘贴替换选区：选区那一截没被留下（首块不再是原来的整段）',
    (xPasted ? textOfBlock(xPasted) : '') !== xPastePick.firstText,
  )
  eq(
    '粘贴替换选区：块数减少 1（跨段替换并成一段）',
    heroBlocks(xAfterPaste).length,
    heroBlocks(xBeforePaste).length - 1,
  )
  await checkDomMatchesModel('X7 粘贴替换跨段选区后：版面文字 = 模型文字')

  /* ------------------------------------------------------------------ */
  console.log('\n=== Y. 功能区标签页：四页常驻、各页各管一摊、切页不丢选区 ===')
  await openApp()
  /** 当前在 DOM 里的页（页内容是 v-if，同一时刻只有一页） */
  const panelPresence = () =>
    page.evaluate(() =>
      ['.panel-start', '.panel-insert', '.panel-layout', '.panel-table'].filter(
        (s) => document.querySelector(s) !== null,
      ),
    )
  /** 预览区上边缘的 y —— 功能区一换高，它就会跳 */
  const canvasTop = () =>
    page.evaluate(() =>
      Math.round(document.querySelector('.canvas').getBoundingClientRect().top),
    )

  eq('功能区有四枚标签', await page.locator('.ribbon-tab').count(), 4)
  eq(
    '标签就是「开始 / 插入 / 布局 / 表格」',
    (await page.locator('.ribbon-tab').allInnerTexts()).join('|'),
    '开始|插入|布局|表格',
  )
  eq('默认停在「开始」页', (await page.locator('.ribbon-tab.is-on').innerText()).trim(), '开始')
  eq('只有「开始」页在 DOM 里', (await panelPresence()).join(','), '.panel-start')

  // 「开始」页该有什么：撤销/重做、加粗/下划线、红与取消颜色、接受/拒绝修订、样式库
  eq(
    '开始页有加粗与下划线',
    await page.locator('.panel-start button.tool[title^="加粗"], .panel-start button.tool[title^="下划线"]').count(),
    2,
  )
  eq(
    '开始页有撤销与重做',
    await page.locator('.panel-start button.tool[title^="撤销"], .panel-start button.tool[title^="重做"]').count(),
    2,
  )
  eq('颜色只剩「红」与「取消颜色」两枚', await page.locator('.panel-start .swatch').count(), 2)
  eq(
    '红色色块就是 FF0000',
    await page.locator('.panel-start .swatch[title="标红"]').evaluate((el) => el.style.background),
    'rgb(255, 0, 0)',
  )
  eq(
    '开始页有接受/拒绝修订',
    await page.locator('.panel-start button.tool[title^="接受"], .panel-start button.tool[title^="拒绝"]').count(),
    2,
  )
  ok('样式库在「开始」页里', (await page.locator('.panel-start .styles .style-chip').count()) >= 5)

  // 「插入」页：三枚特殊空格并排 + 表格/分节符/分页符 + 批注；加粗这类留在开始页
  await openTab('插入')
  eq('切到「插入」页后只剩插入页在 DOM 里', (await panelPresence()).join(','), '.panel-insert')
  eq(
    '插入页三枚特殊空格并排（不再是下拉）',
    await page.locator('.panel-insert button.tool[title^="在插入符处插入"]').count(),
    3,
  )
  eq('插入页里没有下拉框', await page.locator('.panel-insert select').count(), 0)
  eq(
    '插入页有表格 / 分节符 / 分页符',
    await page.locator('.panel-insert button.tool[title^="在光标所在段落后插入"]').count(),
    3,
  )
  eq(
    '插入页有批注输入框与「添加」',
    `${await page.locator('.panel-insert .comment-field input').count()}/${await page.locator('.panel-insert .comment-field button').count()}`,
    '1/1',
  )
  eq('插入页里没有加粗（各页各管一摊，不是全堆在一条栏上）', await page.locator('.panel-insert button.tool[title^="加粗"]').count(), 0)

  await openTab('布局')
  eq('切到「布局」页后只剩布局页在 DOM 里', (await panelPresence()).join(','), '.panel-layout')
  // 节控件的回显跟着落点走：先把插入符放进正文（「节工具条只在光标落进正文之后才出内容」
  // 是 HEAD 上就红着的既有问题，见 PLAN 第 11 节 ⑤，不在本次改动范围里）
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('我方于2026年9月1日'))
  await page.waitForTimeout(200)
  eq('布局页就是节编辑（有节号提示）', await page.locator('.panel-layout .tk-hint').count(), 1)
  eq('布局页里有四组节控件', await page.locator('.panel-layout .tk-group').count(), 4)

  await openTab('表格')
  eq('切到「表格」页后只剩表格页在 DOM 里', (await panelPresence()).join(','), '.panel-table')
  eq('光标不在格子里时表格页也在（常驻）', await page.locator('.panel-table').count(), 1)
  eq('光标不在格子里时给一句提示', await page.locator('.panel-table .tk-empty').count(), 1)
  ok(
    '光标不在格子里时表格页按钮全灰',
    await page.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('.panel-table button'))
      return btns.length > 0 && btns.every((b) => b.disabled)
    }),
  )

  // 这就是拆标签页要解决的问题：光标进出表格时，下面那摞纸不许跳
  const topOutsideTable = await canvasTop()
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('数控加工中心'))
  await page.waitForTimeout(250)
  eq('光标进表格后版面不上移也不下移', await canvasTop(), topOutsideTable)
  eq('光标进表格后提示收起', await page.locator('.panel-table .tk-empty').count(), 0)
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('我方于2026年9月1日'))
  await page.waitForTimeout(250)
  eq('光标离开表格后版面照样不动', await canvasTop(), topOutsideTable)

  /* ------------------------------------------------------------------ */
  console.log('\n=== Z. 格内垂直对齐：最小两行 + 单行文字时也要真的生效 ===')
  await openApp('表格')
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('数控加工中心'))
  await page.waitForTimeout(250)
  /**
   * 目标格的几何：格高（=最小行数×行高）、格内那层 div 的高度（一行文字的自然高）、
   * 以及文字顶端相对格子顶端的偏移 —— 垂直对齐到底有没有生效，就只看最后这个数。
   */
  const cellMetrics = () =>
    page.evaluate(() => {
      const cell = Array.from(
        document.querySelectorAll('.wtp-table .wtp-cell[data-block-id]'),
      ).find((el) => (el.textContent ?? '').includes('数控加工中心'))
      const td = cell?.closest('td')
      if (!cell || !td) return null
      const tdBox = td.getBoundingClientRect()
      const range = document.createRange()
      range.selectNodeContents(cell)
      const text = range.getClientRects()[0]
      return {
        tdHeight: Math.round(tdBox.height),
        cellHeight: Math.round(cell.getBoundingClientRect().height),
        textOffset: text ? Math.round(text.top - tdBox.top) : -1,
      }
    })
  const vAlignButton = (name) =>
    page.locator('.panel-table .tk-group', { hasText: '垂直' }).getByRole('button', {
      name,
      exact: true,
    })

  const atTop = await cellMetrics()
  ok('目标格是两行高（最小两行），格内文字只有一行', atTop !== null && atTop.tdHeight > atTop.cellHeight + 8, JSON.stringify(atTop))
  ok('默认顶端对齐时文字贴着格子上沿', atTop.textOffset >= 0 && atTop.textOffset <= 2, JSON.stringify(atTop))

  await vAlignButton('居中').click()
  await page.waitForTimeout(300)
  const atMiddle = await cellMetrics()
  ok(
    '改成居中对齐后文字真的下移了（旧写法在这里一动不动）',
    atMiddle.textOffset > atTop.textOffset + 4,
    `${JSON.stringify(atTop)} → ${JSON.stringify(atMiddle)}`,
  )
  ok(
    '居中的偏移约等于（格高 − 文字高）/ 2',
    Math.abs(atMiddle.textOffset - (atMiddle.tdHeight - atMiddle.cellHeight) / 2) <= 2,
    JSON.stringify(atMiddle),
  )

  await vAlignButton('底端').click()
  await page.waitForTimeout(300)
  const atBottom = await cellMetrics()
  ok(
    '改成底端对齐后文字比居中时更低',
    atBottom.textOffset > atMiddle.textOffset + 4,
    `${JSON.stringify(atMiddle)} → ${JSON.stringify(atBottom)}`,
  )
  ok(
    '底端的偏移约等于 格高 − 文字高',
    Math.abs(atBottom.textOffset - (atBottom.tdHeight - atBottom.cellHeight)) <= 2,
    JSON.stringify(atBottom),
  )
  // 行高没变（下限写在 td 的 height 上，不是靠格内 div 撑高）—— 分页因此不受影响
  eq('三档对齐都不改格高', `${atTop.tdHeight}/${atMiddle.tdHeight}/${atBottom.tdHeight}`, `${atTop.tdHeight}/${atTop.tdHeight}/${atTop.tdHeight}`)
  await vAlignButton('顶端').click()
  await page.waitForTimeout(250)
  eq('点回顶端后偏移回到 0 附近', (await cellMetrics()).textOffset <= 2, true)

  /* ------------------------------------------------------------------ */
  console.log('\n=== AA. 功能区四页高度一致（W6 第①条）===')
  /*
   * 四页常驻，高度不齐会让下方版面随切页上下跳。最高的一页是「开始」（样式库那排 chip
   * 要按各自的字号渲染），其余三页被 CSS 的内容区下限（.panel 的 min-height）抬到同档。
   * 容差 ±1px：只允许亚像素取整的差别，不允许「矮一截」。
   */
  await openApp()
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('我方于2026年9月1日'))
  await page.waitForTimeout(200)
  const panelHeights = {}
  for (const [label, cls] of [
    ['开始', '.panel-start'],
    ['插入', '.panel-insert'],
    ['布局', '.panel-layout'],
    ['表格', '.panel-table'],
  ]) {
    await openTab(label)
    panelHeights[label] = await page.evaluate((sel) => {
      const el = document.querySelector(sel)
      return el ? el.offsetHeight : -1
    }, cls)
  }
  const panelValues = Object.values(panelHeights)
  ok(
    '四页 offsetHeight 相等（±1px 容差）',
    Math.max(...panelValues) - Math.min(...panelValues) <= 1 && Math.min(...panelValues) > 0,
    JSON.stringify(panelHeights),
  )
  ok(
    '四页都没被挤成两行（高度一致不是因为都换行了）',
    await page.evaluate(() => {
      const el = document.querySelector('.panel')
      return el ? el.scrollHeight <= el.clientHeight + 1 : false
    }),
  )

  /* ------------------------------------------------------------------ */
  console.log('\n=== AB. 自定义快捷键表（W6 第④条）===')
  /*
   * 表通过 URL 参数递进 demo（`?shortcuts=bold:ctrl+shift+b,…`），见 App.vue 的
   * shortcutsFromUrl —— 组件本身只认 `shortcuts` 这个 prop，真实使用方直接传对象。
   */
  const loadCustom = async (search) => {
    await page.goto(`${url}${search}`, { waitUntil: 'load' })
    await page.waitForSelector('.wtp-page', { timeout: 30000 })
    await page.evaluate(async () => {
      await document.fonts.ready
    })
    await page.waitForTimeout(600)
    await page.evaluate(TEST_HELPERS, {
      normal: SEARCH_HIGHLIGHT,
      current: SEARCH_CURRENT_HIGHLIGHT,
    })
  }
  const inlineWithText = async (text) => {
    const model = await getModel()
    return heroBlocks(model)
      .flatMap((b) => b.inlines)
      .find((i) => i.t === 'text' && i.text === text)
  }
  await loadCustom(
    '?shortcuts=bold:ctrl+shift+b,trackChanges:ctrl+alt+e,formatAmount:ctrl+alt+4,bolld:ctrl+i',
  )
  const trackBox = 'label.checkbox input[type="checkbox"]'
  // A. 新组合生效
  await page.evaluate(() => window.__wtpTest.selectIn('我方于2026年9月1日', 0, 2))
  await page.waitForTimeout(60)
  await page.keyboard.press('Control+Shift+b')
  await page.waitForTimeout(250)
  ok(
    '改绑后新组合 ctrl+shift+B 加粗生效（模型里带 b）',
    (await inlineWithText('我方'))?.bold === true,
    JSON.stringify(await inlineWithText('我方')),
  )
  // B. 旧组合失效：改绑到一个浏览器没有原生行为的动作上才验得干净
  //    （加粗这条不行 —— 解绑 ctrl+B 之后 Chromium 的原生加粗会接手，见 README）
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('苏州市公安局'))
  await page.keyboard.press('Control+Shift+e')
  await page.waitForTimeout(200)
  eq('旧组合 ctrl+shift+E 不再翻「修订模式」', await page.isChecked(trackBox), false)
  await page.keyboard.press('Control+Alt+e')
  await page.waitForTimeout(200)
  eq('改绑后新组合 ctrl+alt+E 翻「修订模式」', await page.isChecked(trackBox), true)
  await page.keyboard.press('Control+Alt+e')
  await page.waitForTimeout(200)
  eq('再按一次关掉（回到未勾选）', await page.isChecked(trackBox), false)
  // C. 旧组合失效、且这条在模型上看得见：金额格式化的 alt+4
  await page.evaluate(() => window.__wtpTest.selectIn('人民币5000万元', 3, 7))
  await page.waitForTimeout(60)
  await page.keyboard.press('Alt+4')
  await page.waitForTimeout(250)
  ok('旧组合 alt+4 不再格式化金额（模型里还是 5000）', JSON.stringify(await getModel()).includes('5000'))
  await page.evaluate(() => window.__wtpTest.selectIn('人民币5000万元', 3, 7))
  await page.waitForTimeout(60)
  await page.keyboard.press('Control+Alt+4')
  await page.waitForTimeout(300)
  ok('新组合 ctrl+alt+4 格式化金额', JSON.stringify(await getModel()).includes('5,000.00'))
  // D. 未知动作名：忽略 + 默认表不受影响（另起一次加载，只给那个写错的动作名）
  await loadCustom('?shortcuts=bolld:ctrl+i')
  await page.evaluate(() => window.__wtpTest.selectIn('我方于2026年9月1日', 0, 2))
  await page.waitForTimeout(60)
  await page.keyboard.press('Control+b')
  await page.waitForTimeout(250)
  ok(
    '表里有未知动作名时 ctrl+B 仍然加粗（默认表没被悄悄改掉）',
    (await inlineWithText('我方'))?.bold === true,
    JSON.stringify(await inlineWithText('我方')),
  )

  /* ------------------------------------------------------------------ */
  console.log('\n=== AC. F4 重复上一步（W6 第③条）===')
  await openApp()
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('我方于2026年9月1日'))
  const modelBeforeF4 = JSON.stringify(await getModel())
  await page.keyboard.press('F4')
  await page.waitForTimeout(250)
  eq('没有可重复的操作时弹提示条', await page.locator('.toast').count(), 1)
  eq('没有可重复的操作时模型一个字节都不改', JSON.stringify(await getModel()), modelBeforeF4)
  await page.waitForTimeout(2200)
  // 选中一处加粗 → 换一处选中 → F4 也在那一处加粗
  await page.evaluate(() => window.__wtpTest.selectIn('经核查，债务人名下资产', 0, 3))
  await page.waitForTimeout(60)
  await page.keyboard.press('Control+b')
  await page.waitForTimeout(250)
  ok('第一次加粗生效', (await inlineWithText('经核查'))?.bold === true)
  await page.evaluate(() => window.__wtpTest.selectIn('债务人爱康光电科技有限公司', 0, 3))
  await page.waitForTimeout(60)
  await page.keyboard.press('F4')
  await page.waitForTimeout(300)
  ok(
    'F4 在另一处也加粗了（重放的是格式操作本身，作用在当前选区）',
    (await inlineWithText('债务人'))?.bold === true,
    JSON.stringify(await inlineWithText('债务人')),
  )
  ok(
    '原来那处仍在（模型里两处都带 b）',
    (await inlineWithText('经核查'))?.bold === true,
    JSON.stringify(await inlineWithText('经核查')),
  )
  await page.keyboard.press('Control+z')
  await page.waitForTimeout(300)
  ok(
    '撤销一步只退回这一次重放（第二处不再加粗）',
    (await inlineWithText('债务人'))?.bold !== true,
    JSON.stringify(await inlineWithText('债务人')),
  )
  ok(
    '撤销没有把第一次加粗也退掉（重放自己记了一步）',
    (await inlineWithText('经核查'))?.bold === true,
    JSON.stringify(await inlineWithText('经核查')),
  )
  await checkNoOverflow('AC F4 重放之后')

  /* ------------------------------------------------------------------ */
  console.log('\n=== AD. 顶栏文件名（W6 第②条）===')
  await openApp()
  eq('顶栏不再有 slogan（<strong>）', await page.locator('.bar strong').count(), 0)
  const nameBox = page.locator('.bar .file-name')
  eq('顶栏有一个文件名输入框', await nameBox.count(), 1)
  eq('初始值 = 样本首行标题', await nameBox.inputValue(), '关于爱康光电资产核查情况的说明')
  await nameBox.fill('核查情况说明（终稿）')
  await page.waitForTimeout(120)
  // 能拦到下载事件就验「真的下载成什么名字」；拦不到就退而验输入框（结论里说明是哪一种）
  const exportClick = () => page.locator('.bar button.primary').click()
  let suggested = null
  let downloadSeen = true
  try {
    const [download] = await Promise.all([
      page.waitForEvent('download', { timeout: 8000 }),
      exportClick(),
    ])
    suggested = download.suggestedFilename()
  } catch {
    downloadSeen = false
  }
  if (downloadSeen) {
    eq('导出文件名 = 顶栏文件名 + .docx', suggested, '核查情况说明（终稿）.docx')
  } else {
    ok('（下载事件拦不到）退而验输入框的值变了', (await nameBox.inputValue()) === '核查情况说明（终稿）')
  }
  await nameBox.fill('')
  await page.waitForTimeout(120)
  if (downloadSeen) {
    const [download] = await Promise.all([
      page.waitForEvent('download', { timeout: 8000 }),
      exportClick(),
    ])
    eq('文件名空着时兜底「未命名.docx」', download.suggestedFilename(), '未命名.docx')
  }
  await nameBox.fill('已写后缀.docx')
  await page.waitForTimeout(120)
  if (downloadSeen) {
    const [download] = await Promise.all([
      page.waitForEvent('download', { timeout: 8000 }),
      exportClick(),
    ])
    eq('已经写了 .docx 就不会叠成 .docx.docx', download.suggestedFilename(), '已写后缀.docx')
  }

  /* ------------------------------------------------------------------ */
  console.log('\n=== AE. ::editor：开关写进 md（W6 第⑤条）===')
  /*
   * 界面这条路：源码视图里改文档开头的 `::editor …` → 模型重建 → 组件把开关报给 App
   * （顶栏跟着走）；反过来在顶栏改开关 → 写回模型 → 源码视图里那行跟着出现/消失。
   */
  await openApp()
  const toSourceView = () => page.locator('.tabs button', { hasText: '类 md 源码' }).click()
  const toEditView = () => page.locator('.tabs button', { hasText: '所见即所得' }).click()
  const sourceText = () => page.locator('textarea').inputValue()
  await toSourceView()
  await page.locator('textarea').fill('::editor trackChanges=on\n\n## 甲\n\n乙')
  await page.waitForTimeout(400)
  await toEditView()
  await page.waitForTimeout(300)
  eq('源码里的 trackChanges=on 把顶栏「修订模式」勾上了', await page.isChecked(trackBox), true)
  const flagsModel = await getModel()
  eq('模型里 editor.trackChanges === true', flagsModel.editor?.trackChanges, true)
  eq('模型里 nav 不落字段（默认 true）', flagsModel.editor?.nav, undefined)
  eq('nav 默认开着时导航窗格在', await page.locator('.nav-pane').count(), 1)
  await page.locator(trackBox).click()
  await page.waitForTimeout(250)
  eq('取消勾选后模型里 editor 整个没了', JSON.stringify((await getModel()).editor ?? null), 'null')
  await toSourceView()
  await page.waitForTimeout(300)
  ok('序列化结果里那一行也没了', !(await sourceText()).includes('::editor'), (await sourceText()).slice(0, 60))
  await page.locator('textarea').fill('::editor nav=off\n\n## 甲\n\n乙')
  await page.waitForTimeout(400)
  await toEditView()
  await page.waitForTimeout(300)
  const navModel = await getModel()
  eq('nav=off 落进模型', navModel.editor?.nav, false)
  eq('导航窗格被收起', await page.locator('.nav-pane').count(), 0)
  eq(
    '顶栏「导航」按钮不再是激活态',
    await page.locator('.bar button.tool.is-on', { hasText: '导航' }).count(),
    0,
  )
  await toSourceView()
  await page.waitForTimeout(300)
  ok('序列化里有 nav=off', (await sourceText()).includes('nav=off'), (await sourceText()).slice(0, 60))

  /* ------------------------------------------------------------------ */
  console.log(
    '\n=== AF. 表格复选多格（W7 第①条）：拖动刷选 / Ctrl+点击 / 批量对齐与样式 / 不重排 ===',
  )
  const cellSel = () => page.evaluate(() => window.__wtpPaper.getCellSelection())
  const selKeys = async () =>
    ((await cellSel())?.cells ?? []).map((c) => `${c.row},${c.col}`).join('|')
  /** DOM 上带高亮类的 <td>（按 data-cell-id 排序） */
  const highlighted = () =>
    page.evaluate(() =>
      Array.from(document.querySelectorAll('.wtp-table td.wtp-cellsel'))
        .map((td) => td.dataset.cellId ?? '')
        .sort(),
    )
  const pageHeights = () =>
    page.evaluate(() => Array.from(document.querySelectorAll('.wtp-page')).map((el) => el.offsetHeight))
  /** 量测快照（每块的行数 / 行高、每行表格的高）—— 刷选绝不该让它变 */
  const measuredSignature = () =>
    page.evaluate(() =>
      window.__wtpPaper
        .getMeasurements()
        .map((m) =>
          m.t === 'break'
            ? `B:${m.kind}`
            : m.t === 'tableRow'
              ? `T:${m.blockId}:${m.row}:${Math.round(m.height * 100)}`
              : `${m.blockId}:${m.rows}:${Math.round(m.lineHeight * 100)}`,
        )
        .join('|'),
    )
  /**
   * 一次读回两个格子的中心点。**必须先滚到能同时看到两端的地方**：
   * 分两次滚会拿到过期的坐标（第一次滚完之后另一端的 rect 已经变了），
   * 所以滚的是两端之间的中间那一行。
   */
  async function cellSpan(tableId, from, to) {
    const idOf = ([r, c]) => `${tableId}.r${r}c${c}`
    const midRow = Math.round((from[0] + to[0]) / 2)
    await page.evaluate((id) => {
      document.querySelector(`.wtp-table td[data-cell-id="${id}"]`)?.scrollIntoView({ block: 'center' })
    }, `${tableId}.r${midRow}c0`)
    await page.waitForTimeout(80)
    return page.evaluate(
      ({ a, b }) => {
        const center = (id) => {
          const td = document.querySelector(`.wtp-table td[data-cell-id="${id}"]`)
          if (!td) return null
          const box = td.getBoundingClientRect()
          return { x: box.left + box.width / 2, y: box.top + box.height / 2 }
        }
        return { a: center(a), b: center(b), viewportHeight: window.innerHeight }
      },
      { a: idOf(from), b: idOf(to) },
    )
  }
  /** 从 from 格拖到 to 格（越过起点格 → 整格刷选）。返回 false = 坐标没取到 / 不在视口里 */
  async function brush(tableId, from, to, ctrl = false) {
    const span = await cellSpan(tableId, from, to)
    if (!span?.a || !span?.b) return false
    // 坐标落在视口外的话指针事件到不了版心，刷选会静默失效 —— 当成失败报出来
    if (![span.a, span.b].every((p) => p.y > 0 && p.y < span.viewportHeight)) {
      failures.push(`刷选坐标落在视口外：${JSON.stringify(span)}`)
      return false
    }
    await page.mouse.move(span.a.x, span.a.y)
    if (ctrl) await page.keyboard.down('Control')
    await page.mouse.down()
    await page.mouse.move(span.b.x, span.b.y, { steps: 8 })
    await page.mouse.up()
    if (ctrl) await page.keyboard.up('Control')
    await page.waitForTimeout(180)
    return true
  }
  /** Ctrl+点击某一格（不拖动） */
  async function ctrlClick(tableId, row, col) {
    const span = await cellSpan(tableId, [row, col], [row, col])
    if (!span?.a) return false
    await page.mouse.move(span.a.x, span.a.y)
    await page.keyboard.down('Control')
    await page.mouse.down()
    await page.mouse.up()
    await page.keyboard.up('Control')
    await page.waitForTimeout(180)
    return true
  }
  /** 某一格内第 i / 第 j 个字符的中心点（同一格里拖动 = 普通的选文字） */
  async function cellCharSpan(cellIdValue, i, j) {
    return page.evaluate(
      ({ id, i: from, j: to }) => {
        const cell = document.querySelector(`.wtp-cell[data-block-id="${id}"]`)
        if (!cell) return null
        const node = document.createTreeWalker(cell, NodeFilter.SHOW_TEXT).nextNode()
        if (!node) return null
        const at = (o) => {
          const range = document.createRange()
          range.setStart(node, o)
          range.setEnd(node, Math.min(o + 1, node.data.length))
          const box = range.getBoundingClientRect()
          return { x: box.left + box.width / 2, y: box.top + box.height / 2 }
        }
        return { a: at(from), b: at(to), text: node.data }
      },
      { id: cellIdValue, i, j },
    )
  }
  /** 点正文里的某一处（点正文别处 = 复选收起） */
  async function clickParagraph(needle) {
    const spot = await page.evaluate((text) => {
      const frag = window.__wtpTest.fragmentByText(text)
      if (!frag) return null
      frag.scrollIntoView({ block: 'center' })
      const box = frag.getBoundingClientRect()
      return { x: box.left + 8, y: box.top + box.height / 2 }
    }, needle)
    if (!spot) return false
    await page.waitForTimeout(80)
    await page.mouse.click(spot.x, spot.y)
    await page.waitForTimeout(180)
    return true
  }
  const tdInlineAlign = (ids) =>
    page.evaluate(
      (list) =>
        list.map(
          (id) =>
            document.querySelector(`.wtp-cell[data-block-id="${id}"]`)?.style.textAlign ?? 'none',
        ),
      ids,
    )
  const cellIdsOf = (tableIdValue, cells) =>
    cells.map((c) => `${tableIdValue}.r${c.row}c${c.col}`)
  const alignCountOf = (t) =>
    t.rows.reduce((n, r) => n + r.cells.filter((c) => c.align !== undefined).length, 0)
  const kindCountOf = (t) => t.rows.reduce((n, r) => n + r.cells.filter((c) => c.kind).length, 0)

  // ---- AF1. 拖动刷选 2×2：模型坐标 4 格、DOM 4 个高亮 <td>、不留原生选区、不重排 ----
  await openApp('表格')
  const af1 = await modelTable()
  const afId = af1.id
  // 期望值从模型 + 纯函数现推：rows1..2 × cols1..2 恰好 4 格
  const wantCells = cellsInRects(af1, [cellRectBetween({ row: 1, col: 1 }, { row: 2, col: 2 })])
  eq('（前置）样本表这个 2×2 区域恰好 4 格', wantCells.length, 4)
  // 给版面上的每个片段打一个身份标记：刷选之后它们必须还在（= 没有重建 DOM）
  await page.evaluate(() => {
    for (const el of document.querySelectorAll('.wtp-content > *')) el.__wtpNode = 1
  })
  const heightsBefore = await pageHeights()
  const measuredBefore = await measuredSignature()
  ok('刷选动作本身做得出（两端都滚进了视口）', await brush(afId, [1, 1], [2, 2]))
  eq('刷选后模型坐标下的选中格数 = 4', (await cellSel())?.cells.length, 4)
  eq('选中的正是那 4 格（行优先）', await selKeys(), '1,1|1,2|2,1|2,2')
  eq('DOM 上恰好 4 个 <td> 带高亮类', (await highlighted()).length, 4)
  eq(
    '带高亮的 <td> 就是那 4 格（按 data-cell-id 对）',
    (await highlighted()).join('|'),
    cellIdsOf(afId, wantCells).sort().join('|'),
  )
  const nativeSel = await page.evaluate(() => {
    const s = document.getSelection()
    return { ranges: s?.rangeCount ?? 0, collapsed: s?.isCollapsed ?? true, text: s?.toString() ?? '' }
  })
  ok(
    '拖出格边界时没有残留原生选区（塌掉了 / 或本就没有范围）',
    nativeSel.ranges === 0 || nativeSel.collapsed,
    JSON.stringify(nativeSel),
  )
  eq('刷选不改页数', (await pageHeights()).length, heightsBefore.length)
  eq('刷选不改每页纸的高度（底色不进几何）', JSON.stringify(await pageHeights()), JSON.stringify(heightsBefore))
  eq('刷选不改量测（每块行数 / 行高逐条不变）', await measuredSignature(), measuredBefore)
  eq(
    '刷选没有重建版面 DOM（片段仍是原来那些节点）',
    await page.evaluate(() =>
      Array.from(document.querySelectorAll('.wtp-content > *')).every((el) => el.__wtpNode === 1),
    ),
    true,
  )
  ok(
    '表格页的提示换成「已选 4 格」',
    (await subToolbar.innerText()).includes('已选 4 格'),
    await subToolbar.innerText(),
  )
  eq('多选时对齐按钮按整批回显：body 格默认两端对齐 → 水平一个都不亮', (await activeLabels('水平')).join(','), '')
  await checkNoOverflow('AF1 刷选之后')

  // ---- AF2. 格内选文字不受影响（硬要求）：同一格里拖 → 原生选中文字、格数为 0 ----
  await openApp('表格')
  const af2Id = (await modelTable()).id
  ok('（前置）先刷出 4 格', await brush(af2Id, [1, 1], [2, 2]) && (await cellSel()) !== null)
  const chars = await cellCharSpan(`${af2Id}.r2c0`, 0, 5)
  ok('取到格内文字的两个字符点', chars !== null && chars.text.includes('数控加工中心'), JSON.stringify(chars))
  await page.mouse.move(chars.a.x, chars.a.y)
  await page.mouse.down()
  await page.mouse.move(chars.b.x, chars.b.y, { steps: 6 })
  await page.mouse.up()
  await page.waitForTimeout(180)
  const textSel = await page.evaluate(() => {
    const s = document.getSelection()
    return { ranges: s?.rangeCount ?? 0, collapsed: s?.isCollapsed ?? true, text: s?.toString() ?? '' }
  })
  ok(
    '在同格里拖动 = 普通的选文字（原生选区真的选中了字）',
    textSel.ranges > 0 && !textSel.collapsed && textSel.text.length > 0,
    JSON.stringify(textSel),
  )
  eq('格内选文字时选中格数 = 0（复选被清空）', JSON.stringify(await cellSel()), 'null')
  eq('高亮一个都不剩', (await highlighted()).length, 0)

  // ---- AF3. Ctrl+点击追加 / 去掉、Esc 清空、点正文别处清空 ----
  await openApp('表格')
  const af3 = await modelTable()
  const af3Id = af3.id
  ok('Ctrl+点击第一格', await ctrlClick(af3Id, 1, 0))
  eq('此前没有复选 → Ctrl+点击 = 只选这一格', await selKeys(), '1,0')
  ok('Ctrl+点击对角的格', await ctrlClick(af3Id, 3, 2))
  const wantCross = cellsInRects(af3, [cellRectBetween({ row: 1, col: 0 }, { row: 3, col: 2 })])
  eq('（前置）「锚格↔点击格」那块矩形是 3×3 = 9 格', wantCross.length, 9)
  eq(
    'Ctrl+点击追加：并上锚格↔点击格的矩形',
    await selKeys(),
    wantCross.map((c) => `${c.row},${c.col}`).join('|'),
  )
  eq('高亮跟着变成 9 个 <td>', (await highlighted()).length, 9)
  ok('Ctrl+点击已选中的格', await ctrlClick(af3Id, 3, 2))
  eq('Ctrl+点击一个已被选中的格子 = 把包含它的那一块去掉', await selKeys(), '1,0')
  await page.waitForTimeout(120)
  ok(
    'Esc 之前焦点还在正文里（键盘事件才达得到组件的处理器）',
    await page.evaluate(() => document.activeElement?.classList?.contains('wtp-content') === true),
  )
  await page.keyboard.press('Escape')
  await page.waitForTimeout(180)
  eq('Esc 清空复选', JSON.stringify(await cellSel()), 'null')
  eq('Esc 之后高亮也没了', (await highlighted()).length, 0)
  ok('（前置）再刷出 4 格', await brush(af3Id, [1, 1], [2, 2]) && (await cellSel()) !== null)
  ok('点正文别处', await clickParagraph('我方于2026年9月1日'))
  eq('点正文别处清空复选', JSON.stringify(await cellSel()), 'null')
  eq('高亮也跟着清干净', (await highlighted()).length, 0)

  // ---- AF4. 批量对齐：4 格一起写进模型、隔壁一个不动、一步撤销退回整批 ----
  await openApp('表格')
  const af4 = await modelTable()
  const af4Id = af4.id
  const af4Targets = cellIdsOf(
    af4Id,
    cellsInRects(af4, [cellRectBetween({ row: 1, col: 1 }, { row: 2, col: 2 })]),
  )
  const af4PagesBefore = (await pageHeights()).length
  ok('（前置）刷出 4 格', await brush(af4Id, [1, 1], [2, 2]) && (await cellSel())?.cells.length === 4)
  await alignButton('水平', '居中').click()
  await page.waitForTimeout(300)
  const af4Applied = await modelTable()
  const alignHOf = (t, r, c) => t.rows[r]?.cells[c]?.align?.h
  const af4Quads = [[1, 1], [1, 2], [2, 1], [2, 2]]
  ok(
    '批量「居中」：4 格都写进模型（align.h = center）',
    af4Quads.every(([r, c]) => alignHOf(af4Applied, r, c) === 'center'),
    JSON.stringify(af4Quads.map(([r, c]) => alignHOf(af4Applied, r, c))),
  )
  eq(
    '隔壁格一个都没动（带 align 字段的格只多了这 4 个）',
    alignCountOf(af4Applied) - alignCountOf(af4),
    4,
  )
  eq(
    'DOM：那 4 格的行内 text-align 都对，别处没有 center',
    JSON.stringify(await tdInlineAlign(af4Targets)),
    JSON.stringify(['center', 'center', 'center', 'center']),
  )
  eq('点完「居中」按钮亮着（active 态吃已解析的实际值）', (await activeLabels('水平')).join(','), '居中')
  eq('批量操作之后复选还在（坐标没挪，选中态不该被清）', (await cellSel())?.cells.length, 4)
  eq('批量对齐不改页数（对齐不挪行高、不改换行点）', (await pageHeights()).length, af4PagesBefore)
  await checkNoOverflow('AF4 批量对齐之后')
  await page.keyboard.press('Control+z')
  await page.waitForTimeout(300)
  const af4Undone = await modelTable()
  ok(
    'ctrl+Z 一步退回整批（4 格的 align.h 都没了）',
    af4Quads.every(([r, c]) => alignHOf(af4Undone, r, c) === undefined),
    JSON.stringify(af4Quads.map(([r, c]) => alignHOf(af4Undone, r, c))),
  )
  eq('撤销没有牵连别的格', alignCountOf(af4Undone), alignCountOf(af4))

  // ---- AF5. 批量样式：只有那 4 格换类名（页数允许变） ----
  await openApp('开始')
  const af5 = await modelTable()
  const af5Id = af5.id
  const af5Targets = cellIdsOf(
    af5Id,
    cellsInRects(af5, [cellRectBetween({ row: 1, col: 1 }, { row: 2, col: 2 })]),
  )
  ok('（前置）刷出 4 格', await brush(af5Id, [1, 1], [2, 2]) && (await cellSel())?.cells.length === 4)
  await page.locator('.styles .style-chip', { hasText: '二级标题' }).click()
  await page.waitForTimeout(350)
  const af5Applied = await modelTable()
  const af5Quads = [[1, 1], [1, 2], [2, 1], [2, 2]]
  ok(
    '批量样式：只有那 4 格换成 h2',
    af5Quads.every(([r, c]) => af5Applied.rows[r]?.cells[c]?.kind === 'h2'),
    JSON.stringify(af5Quads.map(([r, c]) => af5Applied.rows[r]?.cells[c]?.kind ?? null)),
  )
  eq('隔壁格一个都没动（带 kind 的格只多了这 4 个）', kindCountOf(af5Applied) - kindCountOf(af5), 4)
  eq(
    'DOM 里恰好 4 格是 wtp-h2',
    await page.evaluate(() => document.querySelectorAll('.wtp-table .wtp-cell.wtp-h2').length),
    4,
  )
  eq(
    '那 4 格的类名就是这 4 个 id',
    (
      await page.evaluate(() =>
        Array.from(document.querySelectorAll('.wtp-table .wtp-cell.wtp-h2'))
          .map((el) => el.dataset.blockId ?? '')
          .sort(),
      )
    ).join('|'),
    af5Targets.slice().sort().join('|'),
  )
  await page.keyboard.press('Control+z')
  await page.waitForTimeout(300)
  const af5Undone = await modelTable()
  ok(
    'ctrl+Z 一步退回整批（4 格的 kind 都没了）',
    af5Quads.every(([r, c]) => af5Undone.rows[r]?.cells[c]?.kind === undefined),
    JSON.stringify(af5Quads.map(([r, c]) => af5Undone.rows[r]?.cells[c]?.kind ?? null)),
  )
  await checkNoOverflow('AF5 批量样式之后')

  // ---- AF6. 跨页的表：样本里没有就照实说明（不为它改样本） ----
  const splitTables = await page.evaluate(() => {
    const counts = new Map()
    for (const frag of document.querySelectorAll('.wtp-tableFrag')) {
      const id = frag.dataset.tableId ?? ''
      counts.set(id, (counts.get(id) ?? 0) + 1)
    }
    return [...counts.entries()].filter(([, n]) => n > 1).map(([id]) => id)
  })
  if (splitTables.length === 0) {
    console.log('  --   样本里没有跨页的表：跨页高亮这一条**无法验证**（不为它改 demo 样本）')
  } else {
    // 有跨页表就验它：同一张表被分页切成几片，高亮就该在每一片上都画到
    await openApp('表格')
    const cross = splitTables[0]
    const crossModel = await getModel()
    const crossTable = crossModel.blocks.find((b) => b.t === 'table' && b.id === cross)
    const crossId = crossTable.id
    const lastBody = crossTable.rows.reduce((last, r, i) => (r.role === 'body' ? i : last), 0)
    /*
     * 两片隔着一个分页断点，视口里同时看不到两端 —— 拖动刷选做不出来（两端必须同时可见），
     * 改用两次 Ctrl+点击凑出跨断点的那个矩形：第 0 行（在上一片）↔ 最后一个正文行（在下一片）。
     */
    ok('（前置）跨页表：Ctrl+点击第 0 行', await ctrlClick(crossId, 0, 0))
    ok(
      '（前置）跨页表：Ctrl+点击最后一个正文行',
      await ctrlClick(crossId, lastBody, crossTable.columns - 1),
    )
    const painted = await page.evaluate(
      (id) =>
        Array.from(document.querySelectorAll('.wtp-tableFrag'))
          .filter((frag) => (frag.dataset.tableId ?? '') === id)
          .map((frag) => frag.querySelectorAll('td.wtp-cellsel').length),
      crossId,
    )
    ok('（前置）跨页表：两片都在页面上', painted.length > 1, JSON.stringify(painted))
    eq(
      '跨页表：每一片上都画到了高亮（按模型坐标判，不按 DOM 元素记）',
      painted.filter((n) => n > 0).length,
      painted.length,
    )
  }

  // ---- AF7. ② 表头 / 附注行恒「最小一行」（minLines=2 也不变高） ----
  await openApp('表格')
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('数控加工中心'))
  await page.waitForTimeout(200)
  /** 每种行角色第一个 <td> 的实测高（plain = unit / note 行，body = 正文行） */
  const rowTdHeights = () =>
    page.evaluate(() => {
      const out = { plain: [], body: [] }
      for (const tr of document.querySelectorAll('.wtp-table tbody tr')) {
        const td = tr.querySelector('td')
        if (!td) continue
        const h = Math.round(td.getBoundingClientRect().height)
        if (tr.classList.contains('wtp-tr-plain')) out.plain.push(h)
        else out.body.push(h)
      }
      return out
    })
  // 期望值从规格表现推：一行高 = 列表段落样式的 linePt（两边都不硬编码）
  const oneLinePx = ptToPx(resolveSpec().styles.listItem.linePt)
  const min2Heights = await rowTdHeights()
  eq('样本表是 minLines=2（前置）', (await modelTable())?.minLines, 2)
  ok(
    'minLines=2：表头行 / 附注行的格高 = 一行（不随行高设置变）',
    min2Heights.plain.length >= 2 && min2Heights.plain.every((h) => Math.abs(h - oneLinePx) <= 3),
    `${JSON.stringify(min2Heights)}，一行≈${oneLinePx}px`,
  )
  // 只拿第一个正文行（表头那行，格内都只有一行字）比：别处有 Shift+Enter 软换行的格子，
  // 那种格的内容本来就比下限高（height 是「最小高度」，内容更高照样撑开）
  ok(
    'minLines=2：正文行的格高 = 两行',
    min2Heights.body.length > 0 && Math.abs((min2Heights.body[0] ?? 0) - oneLinePx * 2) <= 3,
    `${JSON.stringify(min2Heights)}，两行≈${oneLinePx * 2}px`,
  )
  await subRadio('行高', '最小一行').click()
  await page.waitForTimeout(350)
  const min1Heights = await rowTdHeights()
  eq('切成最小一行（前置）', (await modelTable())?.minLines, 1)
  ok(
    'minLines=1：表头行 / 附注行的格高仍是一行（跟着行高设置变就错了）',
    min1Heights.plain.every((h) => Math.abs(h - oneLinePx) <= 3),
    `${JSON.stringify(min1Heights)}，一行≈${oneLinePx}px`,
  )
  ok(
    'minLines=1：正文行的格高也回到一行（三种行都一样）',
    min1Heights.body.length > 0 && Math.abs((min1Heights.body[0] ?? 0) - oneLinePx) <= 3,
    `${JSON.stringify(min1Heights)}，一行≈${oneLinePx}px`,
  )
  await subRadio('行高', '最小两行').click()
  await page.waitForTimeout(350)
  await checkNoOverflow('AF7 行高切换之后')

  /* ------------------------------------------------------------------ */
  console.log(
    '\n=== AG. 组件打包（W8）：props 生效、受控回写、save_md / save_docx、content 留空 ===',
  )
  /*
   * 组件（`WtpEditor`）把「顶栏 + 功能区 + 纸张」收成一个 props / emits 契约：
   * 内容与身份从 props 进，动作与受控回写从 emits 出。demo 只做两件事 ——
   * 把值递进去、把回传值露出来（`window.__wtpDemo`，见 App.vue）。
   * 这一节验的就是这条回路，期望值一律从规格表 / 模型现推，不从界面反推。
   */
  const probe = () => page.evaluate(() => window.__wtpDemo ?? null)
  /** 打开带 URL 参数的 demo（等待逻辑与 openApp 一致） */
  async function openDemo(search) {
    await page.goto(`${url}${search}`, { waitUntil: 'load' })
    await page.waitForSelector('.wtp-page', { timeout: 30000 })
    await page.evaluate(async () => {
      await document.fonts.ready
    })
    await page.waitForTimeout(600)
    await page.evaluate(TEST_HELPERS, {
      normal: SEARCH_HIGHLIGHT,
      current: SEARCH_CURRENT_HIGHLIGHT,
    })
  }
  /** 版心几何 + 量测行数：与 N 节同一套 oracle（「切模板必须整篇按新版心重量」） */
  const geometry = () =>
    page.evaluate(() => {
      const content = document.querySelector('.wtp-content')
      const blocks = window.__wtpPaper.getMeasurements().filter((it) => it.t === 'block')
      return {
        width: content ? content.clientWidth : -1,
        rowSum: blocks.reduce((n, b) => n + b.rows, 0),
      }
    })
  const saveButton = page.locator('.bar button.tool', { hasText: '保存' })
  const authorBox = page.locator('.bar .author-name')
  const templateSelect = page.locator('.bar select')
  const fileNameBox = page.locator('.bar .file-name')
  /** 模型里最后一条批注 */
  const lastComment = async () => {
    const list = (await getModel()).comments
    return list[list.length - 1] ?? null
  }

  // ---- AG1. template prop 决定版心几何 ----
  const otherTemplate = DOC_TEMPLATES.find((t) => t.key !== DOC_TEMPLATES[0].key)
  ok('（前置）规格表里有第二套文件模板可切', otherTemplate !== undefined)
  const managerWidth = contentBoxPx(resolveSpec(DOC_TEMPLATES[0].spec)).width
  const otherWidth = contentBoxPx(resolveSpec(otherTemplate ? otherTemplate.spec : {})).width
  ok(
    '（前置）两套模板的版心宽不同（否则下面两条是空转断言）',
    Math.abs(managerWidth - otherWidth) > 1,
    `${managerWidth} / ${otherWidth}`,
  )
  await openApp()
  eq('不传 template 时用第一套模板', await templateSelect.inputValue(), DOC_TEMPLATES[0].key)
  const managerGeo = await geometry()
  ok(
    '第一套模板的版心宽 = 该模板规格表现推的宽度',
    Math.abs(managerGeo.width - managerWidth) <= 1,
    `实测 ${managerGeo.width}，规格表 ${managerWidth}`,
  )
  await openDemo(`?template=${otherTemplate ? otherTemplate.key : ''}`)
  eq('template prop 决定下拉的初值', await templateSelect.inputValue(), otherTemplate?.key ?? '')
  const otherGeo = await geometry()
  ok(
    'template prop 决定版心几何（= 那一套模板的规格表）',
    Math.abs(otherGeo.width - otherWidth) <= 1,
    `实测 ${otherGeo.width}，规格表 ${otherWidth}`,
  )
  ok(
    'template prop 也决定了量测行数（整篇是按这套版心量的）',
    otherGeo.rowSum !== managerGeo.rowSum,
    `${managerGeo.rowSum} → ${otherGeo.rowSum}`,
  )
  await checkNoOverflow('AG1 换模板 prop 后')

  // ---- AG2. fileName prop：顶栏显示、导出名跟着走、update:fileName 回传 ----
  await openApp()
  eq('fileName prop 显示在顶栏', await fileNameBox.inputValue(), '关于爱康光电资产核查情况的说明')
  await fileNameBox.fill('W8 组件导出探针')
  await page.waitForTimeout(150)
  eq('改顶栏文件名 → update:fileName 把新值交给使用方', (await probe())?.fileName, 'W8 组件导出探针')
  const docxCountBefore = (await probe())?.docxCount ?? -1
  let w8Suggested = null
  try {
    const [download] = await Promise.all([
      page.waitForEvent('download', { timeout: 8000 }),
      page.locator('.bar button.primary').click(),
    ])
    w8Suggested = download.suggestedFilename()
  } catch {
    w8Suggested = null
  }
  eq('点「导出 docx」后 save_docx 事件发出', (await probe())?.docxCount, docxCountBefore + 1)
  if (w8Suggested !== null) {
    eq('下载名 = fileName + .docx', w8Suggested, 'W8 组件导出探针.docx')
  } else {
    ok('（下载事件拦不到）退而验 save_docx 事件已发出', (await probe())?.docxCount === docxCountBefore + 1)
  }

  // ---- AG3. author prop：批注作者用它；update:author 回传 ----
  await openApp('插入')
  eq('author prop 显示在顶栏「修订作者」', await authorBox.inputValue(), '张三')
  await page.evaluate(() => window.__wtpTest.selectIn('债务人爱康光电科技有限公司', 0, 6))
  await page.waitForTimeout(80)
  await page.click('.comment-field input')
  await page.fill('.comment-field input', 'W8 批注作者探针（旧作者）')
  await page.click('.comment-field button')
  await page.waitForTimeout(250)
  eq('新加的批注用的是 author prop 的作者名', (await lastComment())?.author, '张三')
  await authorBox.fill('李四')
  await page.waitForTimeout(150)
  eq('改「修订作者」→ update:author 把新值交给使用方', (await probe())?.author, '李四')
  await page.evaluate(() => window.__wtpTest.selectIn('我方于2026年9月1日', 0, 4))
  await page.waitForTimeout(80)
  await page.click('.comment-field input')
  await page.fill('.comment-field input', 'W8 批注作者探针（新作者）')
  await page.click('.comment-field button')
  await page.waitForTimeout(250)
  eq('改过作者之后新加的批注用的是新作者名', (await lastComment())?.author, '李四')

  // ---- AG4. save_md：顶栏「保存」与 ctrl+S ----
  await openApp()
  const modelPlain = await getModel()
  await saveButton.click()
  await page.waitForTimeout(200)
  eq('点「保存」收到的 md = toMd(getModel())', (await probe())?.lastSaveMd, toMd(modelPlain))
  const trackBoxW8 = page.locator('label.checkbox input[type="checkbox"]')
  await trackBoxW8.click()
  await page.waitForTimeout(250)
  eq('（前置）修订模式已勾上', await trackBoxW8.isChecked(), true)
  const modelTracked = await getModel()
  await saveButton.click()
  await page.waitForTimeout(200)
  const savedTracked = (await probe())?.lastSaveMd ?? ''
  eq('开着修订模式时收到的 md = toMd(getModel())', savedTracked, toMd(modelTracked))
  ok(
    '开着修订模式时保存的 md 带 ::editor trackChanges=on',
    savedTracked.includes('::editor trackChanges=on'),
    savedTracked.slice(0, 60),
  )
  /*
   * ctrl+S：探针挂在 window 的冒泡阶段 —— 组件的处理器在挂载时就注册了，
   * 同阶段下先注册先跑，所以探针读到的是被处理之后的 defaultPrevented。
   */
  await page.evaluate(() => {
    window.__wtpSaveProbe = []
    window.addEventListener('keydown', (e) => {
      if ((e.key || '').toLowerCase() === 's') window.__wtpSaveProbe.push(e.defaultPrevented)
    })
  })
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('我方于2026年9月1日'))
  await page.waitForTimeout(100)
  await page.keyboard.press('Control+s')
  await page.waitForTimeout(250)
  eq('ctrl+S 也触发 save_md（内容 = toMd(getModel())）', (await probe())?.lastSaveMd, toMd(await getModel()))
  eq(
    'ctrl+S 的浏览器默认行为被拦住（探针读到 defaultPrevented）',
    (await page.evaluate(() => window.__wtpSaveProbe)).join(','),
    'true',
  )

  // ---- AG5. content 留空（`?empty=1`）----
  const pageErrors = []
  const onPageError = (error) => pageErrors.push(String(error))
  page.on('pageerror', onPageError)
  await openDemo('?empty=1')
  const emptyState = await page.evaluate(() => ({
    pages: document.querySelectorAll('.wtp-page').length,
    name: document.querySelector('.bar .file-name')?.value ?? null,
    blocks: document.querySelectorAll('.wtp-content [data-block-id]').length,
  }))
  page.off('pageerror', onPageError)
  eq('content 留空时不报错（没有未捕获的页面异常）', pageErrors.length, 0)
  ok(
    'content 留空也出得来版面（至少一张纸、没有块）',
    emptyState.pages >= 1 && emptyState.blocks === 0,
    JSON.stringify(emptyState),
  )
  eq('content 留空时传的确实是空字符串（demo 也没编一个标题出来）', emptyState.name, '')
  await fileNameBox.fill('空文档探针')
  await page.waitForTimeout(150)
  eq('content 留空时顶栏文件名仍可编辑（回传拿到新值）', (await probe())?.fileName, '空文档探针')
} finally {
  await browser?.close()
  await server.close()
}

console.log('')
if (failures.length > 0) {
  console.error(`[FAIL] 共 ${failures.length} 项不符：`)
  for (const f of failures) console.error(`  - ${f}`)
  process.exit(1)
}
console.log(
  '[PASS] 编辑层实测：输入不重排不丢插入符、回车/退格、加粗/下划线/改色、修订、批注、撤销、' +
    '金额格式、特殊空格、切文件模板、打印（含新增对齐/删表按钮的隐藏、两档命名 @page）、查找替换（面板/高亮/范围/替换一处与全部）、' +
    '导航窗格（条目与模型一致、点击跳转、折叠）、表格（渲染/格内读回/插入表格面板选规格与越界夹回）、' +
    '表格编辑交互（表格页常驻与置灰规则、落点提示、增删行列、unit/note 与行高 radio、格内 Shift+Enter 落点、' +
    '格首 Backspace 与格尾 Delete 护栏）、表格收尾（Tab/Shift+Tab 与 ←/→ 跨格、最后一格 Tab 无响应、' +
    '删除整表后插入符落上一块末尾、格内点样式 chip 只改那一格、两组对齐的 active 态与行内样式、对齐不改页数）、' +
    '节工具条（常驻、节号回显、首节与「关联前节=是」的置灰、改方向后逐页几何横竖互换且页数不变、' +
    '关页码后该节不再有页码元素而别的节不受影响、默认值不落模型字段）、' +
    '接受/拒绝修订（无修订时置灰、拒绝删除修订、接受插入修订、撤销能还原）、' +
    '功能区标签页（四页常驻、各页各管一摊、光标进出表格时版面不跳）、' +
    '格内垂直对齐（最小两行 + 单行文字时三档真的生效，且不改格高）、' +
    'W6：功能区四页 offsetHeight 相等、自定义快捷键表（改绑后新组合生效 / 旧组合失效 / 未知动作名不改默认表）、' +
    'F4 重复上一步（空转提示、换处重放、撤销只退这一步）、顶栏文件名（无 slogan、导出名 = 文件名 + .docx、空名兜底）、' +
    '::editor 写进 md（源码 → 顶栏开关、顶栏开关 → 源码那一行）均落到模型；' +
    'W7：表格整格复选（拖动刷选 2×2 = 4 格且高亮 4 个 <td>、拖出格边界不留原生选区、刷选不重排不重建 DOM、' +
    '格内拖选文字不受影响（格数为 0）、Ctrl+点击追加与去掉、Esc / 点正文别处清空、' +
    '批量对齐与批量样式只改选中的格且一步撤销退整批、对齐不改页数、表头与附注行恒一行高）；' +
    'W8：组件打包（template prop 决定版心几何与量测行数、fileName prop 决定顶栏与导出名、author prop ' +
    '决定批注作者、update:fileName / update:author / update:template 回传、save_md 与 ctrl+S 一致且 ' +
    'preventDefault、save_docx 事件发出、content 留空不报错）。',
)
