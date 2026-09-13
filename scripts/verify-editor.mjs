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
 *   6. 切文件模板：页数变了、插入符与选区都按坐标找回、文字没丢。
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
  computeNumbering,
  resolveSpec,
} from '../dist-lib/wordtohtml.mjs'

const here = dirname(fileURLToPath(import.meta.url))
const root = resolve(here, '..')
const EDGE = 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe'
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

async function openApp() {
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
}

const getModel = () => page.evaluate(() => window.__wtpPaper.getModel())
const heroBlocks = (model) => model.blocks.filter((b) => b.t === 'textBlock')
const textOfBlock = (block) =>
  block.inlines
    .filter((i) => i.t === 'text')
    .map((i) => i.text)
    .join('')

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
  browser = await chromium.launch({ executablePath: EDGE, headless: true })
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
  await page.click('.swatch[title="红"]')
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

  /* ------------------------------------------------------------------ */
  console.log('\n=== H. 批注：选中 → 写内容 → 侧栏出现 ===')
  await openApp()
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
  await openApp()
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
  await openApp()
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
  await openApp()
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
  await openApp()
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
  eq('换页只推进一页 —— 与 Word 实测一致', both.pages, beforeBoth + 1)
  await checkNoOverflow('H5 两枚标记连在一起后')

  /* ------------------------------------------------------------------ */
  console.log('\n=== H6. 文末插分节符：Word 会多留一张空白页 ===')
  await openApp()
  const beforeTail = await page.evaluate(() => document.querySelectorAll('.wtp-page').length)
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
  eq('文末分节符多留了一张空白页', tail.pages, beforeTail + 1)
  eq('最后一张确实没有片段', tail.lastFragments, 0)
  eq('空白页页码重排为 1', tail.lastNumber, '1')
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
  console.log('\n=== M. 工具栏：插入三种特殊空格（模型里是确切码点） ===')
  await openApp()
  /** 走真实的下拉 change（焦点会离开正文，靠选区/落点记录回退） */
  async function pickSpace(kind) {
    await page.selectOption('.toolbar select', kind)
    await page.waitForTimeout(250)
  }
  /** 直接在正文里改下拉值触发 change（焦点不动，走实时选区那条路） */
  async function fireSpaceChange(kind) {
    await page.evaluate((k) => {
      const sel = document.querySelector('.toolbar select')
      sel.value = k
      sel.dispatchEvent(new Event('change', { bubbles: true }))
    }, kind)
    await page.waitForTimeout(250)
  }

  await page.evaluate(() => window.__wtpTest.caretAtEndOf('苏州市公安局'))
  await fireSpaceChange('em')
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

  // 另外两个：半宽与四分之一宽。这次走真实下拉（焦点被拿走，考的是选区/落点记录的兜底）
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('苏州市公安局'))
  await pickSpace('en')
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('苏州市公安局'))
  await pickSpace('quarterEm')
  const spacesModel = await getModel()
  const spacesBlock = heroBlocks(spacesModel).find((b) => textOfBlock(b).includes('\u2002'))
  const points = spacesBlock ? Array.from(textOfBlock(spacesBlock)).map((c) => c.codePointAt(0)) : []
  ok('模型里有 U+2002', points.includes(0x2002), JSON.stringify(points))
  ok('模型里有 U+2005', points.includes(0x2005), JSON.stringify(points))
  ok('三种空格都插在段末，顺序与操作一致', spacesBlock
    ? textOfBlock(spacesBlock).endsWith('\u2003\u2002\u2005')
    : false, JSON.stringify(textOfBlock(spacesBlock)))
  ok('普通空格没有混进来', !points.includes(0x20), JSON.stringify(points))

  // 下拉必须复位回占位项，否则再选同一项不会再触发 change
  eq('下拉复位回占位项', await page.evaluate(() => document.querySelector('.toolbar select').value), '')

  // 有选区时替换选区（与 insertText / replaceRange 的既有语义一致）
  await page.evaluate(() => window.__wtpTest.selectIn('我方于2026年9月1日', 0, 2))
  await page.waitForTimeout(60)
  await fireSpaceChange('en')
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
  await openApp()
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

  await page.selectOption('.bar select', otherKey)
  await page.waitForTimeout(1000)

  const pagesAfter = await page.evaluate(() => document.querySelectorAll('.wtp-page').length)
  const caretAfter = await page.evaluate(() => window.__wtpTest.caretInfo())
  ok('切模板后页数变了（说明整篇按新版心重量了）', pagesAfter !== pagesBefore, `${pagesBefore} → ${pagesAfter}`)
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
  await openApp()
  // 打开查找面板：打印时它和导航窗格都必须一起消失。
  // 落点放进表格格子里 —— 上下文工具条（.sub-toolbar）只在格内出现，打印时它也必须一起消失。
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('数控加工中心'))
  await page.keyboard.press('Control+f')
  await page.waitForSelector('.search-panel', { timeout: 3000 })
  // 插入表格面板也要一起验：它和查找面板一样是浮层，打印时同样不该出现
  await page.locator('button.tool[title^="在光标所在段落后插入一张空表格"]').click()
  await page.waitForSelector('.table-panel', { timeout: 3000 })

  /** 打印媒体下这些选择器的 display（元素不存在时给 'missing'，别把「没有」当成「隐藏了」） */
  const printDisplay = () =>
    page.evaluate(() => {
      const sels = [
        '.bar',
        '.styles',
        '.toolbar',
        '.sub-toolbar',
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
    '.styles',
    '.toolbar',
    '.sub-toolbar',
    '.wtp-comments',
    '.wtp-measure-root',
    '.wtp-break',
    '.nav-pane',
    '.search-panel',
    '.table-panel',
  ]) {
    eq(`打印时隐藏 ${sel}`, printEdit[sel], 'none')
  }
  eq('第一张纸前面不再断页', printEdit.firstBreak, 'auto')
  eq('后续每张纸都在新的一页开始', printEdit.secondBreak, 'page')
  eq('打印时纸张不带屏幕上的阴影', printEdit.boxShadow, 'none')
  eq('打印时版面页数不变', printEdit.pages, await page.evaluate(() => document.querySelectorAll('.wtp-page').length))

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
        cells: document.querySelectorAll('.wtp-table .wtp-listItem[data-block-id]').length,
        firstCellText: document.querySelector('.wtp-table .wtp-listItem[data-block-id]')
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
  await openApp()
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
  const subToolbar = page.locator('.sub-toolbar')
  const subButton = (name) => subToolbar.getByRole('button', { name, exact: true })
  /** 上下文工具条里某个 radio 组里的一个选项（组按左边的标签文字定位，避免「有/无」重名） */
  const subRadio = (group, name) =>
    subToolbar.locator('.tk-group', { hasText: group }).getByRole('radio', { name, exact: true })

  // ---- V1. 工具条的出现/消失 + 下方插入行 ----
  await openApp()
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('数控加工中心'))
  await page.waitForTimeout(200)
  eq('光标进表格后上下文工具条出现', await subToolbar.count(), 1)
  ok(
    '工具条显示落点（行/列都是下标，显示时 +1）',
    (await subToolbar.innerText()).includes('第 3 行第 1 列'),
    await subToolbar.innerText(),
  )
  await page.evaluate(() => window.__wtpTest.caretAtEndOf('我方于2026年9月1日'))
  await page.waitForTimeout(200)
  eq('光标离开表格后工具条消失', await subToolbar.count(), 0)

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
  await openApp()
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
  await openApp()
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
  await openApp()
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
  const typed = await modelTable()
  const typedInlines = cellInlines(typed, shiftRow, 0)
  const breakIdx = typedInlines.findIndex((i) => i.t === 'break')
  const typedIdx = typedInlines.findIndex((i) => i.t === 'text' && i.text.includes('X'))
  ok('Shift+Enter 后插入符落在换行之后（敲的字排在 break 之后）', breakIdx >= 0 && typedIdx > breakIdx)
  eq('换行后敲的字没打回上一行', cellText(typed, shiftRow, 0), '检测仪器X')
  const v4Caret = await page.evaluate(() => window.__wtpTest.caretInfo())
  eq('插入符仍在同一格', v4Caret?.blockId, `${typed?.id ?? ''}.r${shiftRow}c0`)
  await checkNoOverflow('V4 Shift+Enter 后')

  // ---- V5. 边界护栏：格首 Backspace、格尾 Delete 都不许动模型/DOM ----
  await openApp()
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
  const afterDelete = await modelTable()
  eq('格尾 Delete：行数不变', afterDelete?.rows.length, guardTable?.rows.length)
  eq('格尾 Delete：格内文字不变', cellText(afterDelete, 4, 0), '检测仪器')
  eq(
    '格尾 Delete：<td> 数没变',
    (await tableDom()).tds,
    guardDom.tds,
  )
  await checkNoOverflow('V5 边界护栏后')
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
    '金额格式、特殊空格、切文件模板、打印、查找替换（面板/高亮/范围/替换一处与全部）、' +
    '导航窗格（条目与模型一致、点击跳转、折叠）、表格（渲染/格内读回/插入表格面板选规格与越界夹回）、' +
    '表格编辑交互（上下文工具条与落点提示、增删行列、unit/note 与行高 radio、格内 Shift+Enter 落点、' +
    '格首 Backspace 与格尾 Delete 护栏）均落到模型。',
)
