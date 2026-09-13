/**
 * P2 的浏览器侧验收：用真实浏览器打开 demo，实测分页结果、排版样式与双页并排。
 *
 * 为什么必须做这一步：分页器本身的单测只证明「给定量测值能算对」，
 * 而「量得对不对、渲染出来和 docx 像不像」只有真排一遍才知道。
 * 这里断言的是四类事实：
 *   1. 每一页的内容都没有超出该页版心（超了就说明装箱算术和真实布局脱节）；
 *   2. 每套文件模板各自的字号/行高/字重/对齐/缩进/段距、版心宽，与规格表一致；
 *   3. 分节后页码确实重排，且两套模板的页数不同（切模板会清量测缓存、整篇重量，
 *      页数不动就说明还在拿旧版心的量测值算页码）；
 *   4. 双页并排：宽视口下确实有两页同处一行且纸宽未被压缩，窄视口回落成一页一排。
 *
 * 用 Vite 的编程式 API 起 dev server（不用子进程，也就不存在残留进程问题），
 * 走系统 Edge（playwright-core 不自带浏览器，也不下载）。
 *
 * 用法：node scripts/verify-browser.mjs
 */

import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

import { chromium } from 'playwright-core'
import { createServer } from 'vite'

import {
  DOC_TEMPLATES,
  STYLE_KEYS,
  commentScopes,
  contentBoxPx,
  lengthToPx,
  lineSpacePt,
  parseMd,
  ptToPx,
  resolveSections,
  resolveSpec,
  toMd,
} from '../dist-lib/wordtohtml.mjs'

const here = dirname(fileURLToPath(import.meta.url))
const root = resolve(here, '..')
const EDGE = 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe'
const PORT = 5199

const failures = []

function eq(label, actual, expected) {
  if (actual === expected) return true
  failures.push(`${label} → 期望 ${JSON.stringify(expected)}，实际 ${JSON.stringify(actual)}`)
  return false
}

function approx(label, actual, expected, tol = 0.6) {
  if (Number.isFinite(actual) && Math.abs(actual - expected) <= tol) return true
  failures.push(`${label} → 期望 ≈${Math.round(expected * 100) / 100}（容差 ${tol}），实际 ${actual}`)
  return false
}

function ok(label, condition, detail = '') {
  if (condition) return true
  failures.push(`${label}${detail ? ` — ${detail}` : ''}`)
  return false
}

const round2 = (v) => Math.round(v * 100) / 100

/**
 * 页面里跑的一遍实测。返回的东西全部是「浏览器亲眼看到的」：
 * 每页的占用与溢出、每页版心的实际宽度、每种样式的计算值、量测容器给的行盒。
 * 期望值一律不在页面里算 —— 那是 node 侧按 resolveSpec() 推导的事。
 */
const analyzePage = (kinds) => {
  const round = (v) => Math.round(v * 1000) / 1000

  const pages = Array.from(document.querySelectorAll('.wtp-page')).map((pageEl, index) => {
    const content = pageEl.querySelector('.wtp-content')
    const kids = content ? Array.from(content.children) : []
    const rect = pageEl.getBoundingClientRect()
    let used = 0
    const items = kids.map((el, i) => {
      const cs = getComputedStyle(el)
      const height = round(el.getBoundingClientRect().height)
      const mt = parseFloat(cs.marginTop) || 0
      const mb = i < kids.length - 1 ? parseFloat(cs.marginBottom) || 0 : 0
      used += height + mt + mb
      return {
        kind: el.className.replace('wtp-', ''),
        blockId: el.dataset.blockId ?? '',
        // 表格片段的外层按设计不挂 data-block-id（格内 div 才是可寻址的块），
        // 它的身份在 data-table-id 上 —— 对账表格时必须用这个键，否则永远取不到渲染高
        tableId: el.dataset.tableId ?? '',
        height,
        marginTop: round(mt),
        marginBottom: round(mb),
        text: (el.textContent ?? '').slice(0, 14),
      }
    })
    return {
      index: index + 1,
      number: (pageEl.querySelector('.wtp-page-number')?.textContent ?? '').trim(),
      contentHeight: content ? round(content.clientHeight) : 0,
      contentWidth: content ? round(content.clientWidth) : 0,
      usedHeight: round(used),
      fragments: kids.length,
      items,
      rect: { top: round(rect.top), left: round(rect.left), width: round(rect.width) },
      first: (kids[0]?.textContent ?? '').slice(0, 18),
      last: (kids[kids.length - 1]?.textContent ?? '').slice(-18),
    }
  })

  // 样式必须在「挂到 body 上的探针元素」上读，不能读分页片段本身：
  // 页首那一块会被刻意去掉段前距，读它会拿到豁免后的值，从而掩盖规格表的问题。
  const styles = {}
  for (const kind of kinds) {
    const el = document.createElement('div')
    el.className = `wtp-${kind}`
    el.textContent = '测量'
    document.body.appendChild(el)
    const cs = getComputedStyle(el)
    styles[kind] = {
      fontFamily: cs.fontFamily,
      fontSize: parseFloat(cs.fontSize),
      fontWeight: cs.fontWeight,
      textAlign: cs.textAlign,
      lineHeight: parseFloat(cs.lineHeight),
      textIndent: parseFloat(cs.textIndent),
      marginTop: parseFloat(cs.marginTop),
      marginBottom: parseFloat(cs.marginBottom),
    }
    el.remove()
  }

  // 续排片段只取段落片段（表格片段的外层没有 data-block-id，它没有「首行缩进」这回事）
  const continuations = Array.from(
    document.querySelectorAll('[data-continuation="1"][data-block-id]'),
  ).map((el) => {
    const cs = getComputedStyle(el)
    return {
      textIndent: parseFloat(cs.textIndent),
      marginTop: parseFloat(cs.marginTop),
      text: (el.textContent ?? '').slice(0, 16),
    }
  })

  // 页带的可用宽与实际内容宽：并排时「不横向溢出」就看这一对数
  const strip = document.querySelector('.wtp-pages')
  const textarea = document.querySelector('textarea')
  // 编辑层成了主界面以后，源码不再从 textarea 反推 —— 直接从组件拿模型，
  // 它才是「导出 docx 的依据」。textarea 只在源码模式下存在，取到就顺带带上。
  const model = window.__wtpPaper?.getModel?.() ?? null
  // 批注：侧栏条目 + 正文锚点。侧栏内容必须与正文锚点对得上，
  // 否则「看得到底色、看不到内容」那个缺陷会以另一种形式回来。
  const commentAside = document.querySelector('.wtp-comments')
  const commentSidebar = {
    present: commentAside !== null,
    items: commentAside
      ? Array.from(commentAside.querySelectorAll('li .wtp-comment-item')).map((b) => ({
          scope: (b.querySelector('.wtp-comment-scope')?.textContent ?? '').trim(),
          text: (b.querySelector('.wtp-comment-text')?.textContent ?? '').trim(),
        }))
      : [],
  }
  const commentAnchors = Array.from(document.querySelectorAll('.wtp-comment')).map((el) => ({
    id: el.dataset.comment ?? '',
    text: el.textContent ?? '',
  }))

  return {
    pageCount: pages.length,
    pages,
    strip: strip
      ? { clientWidth: round(strip.clientWidth), scrollWidth: round(strip.scrollWidth) }
      : null,
    styles,
    continuations,
    commentSidebar,
    commentAnchors,
    model,
    source: textarea ? textarea.value : '',
  }
}

/** 选中某套文件模板并等重排稳定：先等版心宽对上，再等页数不再变 */
async function selectTemplate(page, template) {
  const spec = resolveSpec(template.spec)
  const want = contentBoxPx(spec).width
  await page.selectOption('.bar select', template.key)
  await page.waitForFunction(
    (w) => {
      const el = document.querySelector('.wtp-content')
      return el !== null && Math.abs(el.clientWidth - w) <= 1
    },
    want,
    { timeout: 15000 },
  )
  let last = -1
  for (let i = 0; i < 30; i += 1) {
    const n = await page.evaluate(() => document.querySelectorAll('.wtp-page').length)
    if (n === last && n > 0) return
    last = n
    await page.waitForTimeout(120)
  }
}

/**
 * 逐节信息（从模型现推）。verify-page-count 拿它当 oracle：
 * 「重排的节数」= 除首节外声明了「从 1 开始」的节数，所以预览里页码 1 的出现次数
 * 应当恰好是 `1 + restartCount`（首节的第一页也是 1）。
 * 从模型推而不是从预览数自己，才能真的验到东西。
 */
function sectionInfo(report, template) {
  const model = report.model ?? parseMd(report.source ?? '')
  const lives = resolveSections(model, resolveSpec(template.spec))
  return {
    sectionCount: lives.length,
    restartCount: lives.filter((s) => s.settings.restartAtOne).length,
  }
}

const server = await createServer({
  root,
  logLevel: 'warn',
  server: { port: PORT, strictPort: true },
})
await server.listen()
const url = server.resolvedUrls?.local?.[0] ?? `http://localhost:${PORT}/`
console.log(`dev server: ${url}`)

let browser
try {
  browser = await chromium.launch({ executablePath: EDGE, headless: true })
  const page = await browser.newPage({
    viewport: { width: 1700, height: 1100 },
    deviceScaleFactor: 2,
  })

  await page.goto(url, { waitUntil: 'load' })
  await page.waitForSelector('.wtp-page', { timeout: 30000 })
  // 等字体就绪：组件就绪后会再排一次，此时量到的才是最终结果
  await page.evaluate(async () => {
    await document.fonts.ready
  })
  await page.waitForTimeout(800)

  /* ------------------------- 每套模板各跑一遍 ------------------------- */

  const seen = []
  for (const template of DOC_TEMPLATES) {
    const spec = resolveSpec(template.spec)
    await selectTemplate(page, template)
    const report = await page.evaluate(analyzePage, STYLE_KEYS)
    seen.push({ template, spec, report })
    const tag = `「${template.label}」`

    console.log(`\n\n######## ${tag}（${template.key}，版心 ${round2(contentBoxPx(spec).width)}px）########`)

    console.log('\n=== 1. 分页（每页内容不得超出该页版心）===')
    ok(`${tag} 至少渲染出一页`, report.pageCount > 0)
    for (const p of report.pages) {
      ok(
        `${tag} 第${p.index}页未溢出`,
        p.usedHeight <= p.contentHeight + 1,
        `已用 ${p.usedHeight}px / 版心 ${p.contentHeight}px`,
      )
      ok(`${tag} 第${p.index}页有页码`, p.number !== '')
      console.log(
        `  ${p.usedHeight <= p.contentHeight + 1 ? 'ok  ' : 'FAIL'} 第${p.index}页 ` +
          `页码"${p.number}" 片段${p.fragments} 占位${p.usedHeight}/${p.contentHeight}px ` +
          `剩余${round2(p.contentHeight - p.usedHeight)}px`,
      )
      for (const it of p.items) {
        console.log(
          `        ${it.kind.padEnd(10)} 高${String(round2(it.height)).padStart(7)} ` +
            `前${String(round2(it.marginTop)).padStart(5)} 后${String(round2(it.marginBottom)).padStart(5)} ` +
            `"${it.text}"`,
        )
      }
    }
    const numbers = report.pages.map((p) => p.number).join(',')
    console.log(`  页码序列：${numbers}`)
    // W5：页码「重排」由**该节的设置**决定（新插的分节符默认不重排），所以不能写死
    // 「末页是 1」。能断言的是两条硬事实：页码 1 的出现次数 = 首节 + 声明了「从 1 开始」
    // 的节数；且序列里确实出现过回退（一节都不重排的话它必然单调不减）。
    const restartInfo = sectionInfo(report, template)
    const ones = report.pages.filter((p) => p.number === '1').length
    eq(
      `${tag} 页码 1 的出现次数（首节 1 次 + 声明重排的 ${restartInfo.restartCount} 次）`,
      ones,
      1 + restartInfo.restartCount,
    )
    ok(
      `${tag} 页码确实在某节处重排过（序列出现回退）`,
      report.pages.some((p, i) => i > 0 && Number(p.number) <= Number(report.pages[i - 1]?.number)),
      `页码序列 ${numbers}`,
    )

    console.log('\n=== 2. 预览样式与版心宽 vs 规格表 ===')
    for (const kind of Object.keys(spec.styles)) {
      const s = spec.styles[kind]
      const got = report.styles[kind]
      if (!got) {
        failures.push(`预览里找不到 .wtp-${kind} 元素`)
        continue
      }
      const checks = [
        approx(`${tag} ${kind}·字号`, got.fontSize, ptToPx(s.sizePt)),
        s.lineRule === 'auto' ? true : approx(`${tag} ${kind}·行高`, got.lineHeight, ptToPx(s.linePt)),
        eq(`${tag} ${kind}·字重`, got.fontWeight, s.bold ? '700' : '400'),
        eq(`${tag} ${kind}·对齐`, got.textAlign, s.align === 'both' ? 'justify' : s.align),
        approx(`${tag} ${kind}·首行缩进`, got.textIndent, s.firstLineChars * ptToPx(s.sizePt)),
        approx(`${tag} ${kind}·段前`, got.marginTop, ptToPx(lineSpacePt(s.spaceBeforeLines, spec))),
        approx(`${tag} ${kind}·段后`, got.marginBottom, ptToPx(lineSpacePt(s.spaceAfterLines, spec))),
        ok(
          `${tag} ${kind}·字体栈含中西文`,
          got.fontFamily.includes(s.ascii) && got.fontFamily.includes(s.eastAsia),
          got.fontFamily,
        ),
      ]
      console.log(
        `${checks.every(Boolean) ? 'ok  ' : 'FAIL'} ${kind.padEnd(10)} ` +
          `${Math.round(got.fontSize * 100) / 100}px/` +
          `${Number.isFinite(got.lineHeight) ? Math.round(got.lineHeight * 100) / 100 : 'normal'}px ` +
          `字重${got.fontWeight} 对齐${got.textAlign} 缩进${Math.round(got.textIndent)}px ` +
          `段前后${Math.round(got.marginTop)}/${Math.round(got.marginBottom)}px`,
      )
    }
    // 版心宽是分页算术的输入，必须与规格表逐像素对上：纸被 flex 压窄就全错
    const wantContent = contentBoxPx(spec)
    for (const p of report.pages) {
      approx(`${tag} 第${p.index}页版心宽`, p.contentWidth, wantContent.width, 1)
    }
    console.log(
      `  ok   版心宽 ${report.pages[0]?.contentWidth}px（规格表 ${round2(wantContent.width)}px），` +
        `版心高 ${report.pages[0]?.contentHeight}px（规格表 ${round2(wantContent.height)}px）`,
    )

    console.log('\n=== 3. 页首与续排的间距豁免 ===')
    for (const p of report.pages) {
      // 空白页（文末分节符留下的那种）没有片段，谈不上「首块段前距」
      if (p.items.length === 0) {
        ok(`${tag} 第${p.index}页是空白页（Word 也会留这一页）`, p.fragments === 0)
        continue
      }
      ok(
        `${tag} 第${p.index}页首块段前距为 0`,
        (p.items[0]?.marginTop ?? -1) === 0,
        `实际 ${p.items[0]?.marginTop}`,
      )
    }
    ok(`${tag} 样本里出现了跨页续排`, report.continuations.length > 0)
    for (const c of report.continuations) {
      ok(`${tag} 续排块取消首行缩进「${c.text}」`, c.textIndent === 0, `实际 ${c.textIndent}px`)
      ok(`${tag} 续排块段前距为 0「${c.text}」`, c.marginTop === 0, `实际 ${c.marginTop}px`)
    }
    console.log(
      `  ok   续排片段 ${report.continuations.length} 处，页首块 ${report.pages.length} 处`,
    )

    console.log('\n=== 4. 量测值 vs 渲染值（分页算术的对账）===')
    const measurements = await page.evaluate(() => {
      const paper = window.__wtpPaper
      if (!paper) return null
      return paper.getMeasurements().map((m) =>
        m.t === 'break'
          ? { t: 'break', kind: m.kind, id: m.blockId }
          : m.t === 'tableRow'
            ? { t: 'tableRow', id: m.blockId, row: m.row, height: m.height }
            : {
                t: 'block',
                id: m.blockId,
                kind: m.kind,
                rows: m.rows,
                lineHeight: m.lineHeight,
                before: m.spaceBefore,
                after: m.spaceAfter,
                length: m.displayLength,
              },
      )
    })

    if (!measurements) {
      failures.push('拿不到量测值：window.__wtpPaper 未挂载（demo 的 dev 钩子失效？）')
      console.log('  FAIL 无法读取量测值')
    } else {
      const rendered = new Map()
      for (const p of report.pages) {
        for (const it of p.items) {
          const key = it.tableId || it.blockId
          if (key === '') continue
          const cur = rendered.get(key) ?? { height: 0, pieces: 0 }
          cur.height += it.height
          cur.pieces += 1
          rendered.set(key, cur)
        }
      }

      for (const m of measurements) {
        if (m.t === 'break') {
          console.log(`        [${m.kind === 'section' ? '分节符' : '分页符'}]`)
          continue
        }
        if (m.t === 'tableRow') continue
        const r = rendered.get(m.id) ?? { height: 0, pieces: 0 }
        const fromMeasurement = round2(m.rows * m.lineHeight)
        const match = Math.abs(fromMeasurement - r.height) < 1.5
        if (!match) {
          failures.push(
            `${tag} ${m.kind} 量测 ${m.rows}行×${round2(m.lineHeight)}=${fromMeasurement}px，` +
              `渲染 ${r.height}px`,
          )
        }
        console.log(
          `  ${match ? 'ok  ' : 'BAD '} ${m.kind.padEnd(10)} ` +
            `量测 ${m.rows}行×${String(round2(m.lineHeight)).padStart(6)}=${String(fromMeasurement).padStart(7)}px ` +
            `渲染${String(r.height).padStart(7)}px(${r.pieces}片) 段前后${round2(m.before)}/${round2(m.after)}`,
        )
      }

      // 表格：按表把各行量测高加起来，与渲染出来的表格片高对账
      //（同一张表可能被分页切成几片，渲染那侧也按 blockId 累加，所以两边口径一致）
      const tableTotals = new Map()
      for (const m of measurements) {
        if (m.t !== 'tableRow') continue
        tableTotals.set(m.id, (tableTotals.get(m.id) ?? 0) + m.height)
      }
      for (const [id, measuredTotal] of tableTotals) {
        const r = rendered.get(id) ?? { height: 0, pieces: 0 }
        const match = Math.abs(measuredTotal - r.height) < 1.5
        if (!match) {
          failures.push(
            `${tag} 表格 量测各行合计 ${round2(measuredTotal)}px，渲染 ${r.height}px`,
          )
        }
        console.log(
          `  ${match ? 'ok  ' : 'BAD '} 表格       ` +
            `量测各行合计=${String(round2(measuredTotal)).padStart(7)}px ` +
            `渲染${String(r.height).padStart(7)}px(${r.pieces}片) id=${id}`,
        )
      }
    }
  }

  const [first, second] = seen
  const firstReport = first.report

  console.log('\n=== 5. 批注侧栏（侧栏、正文锚点、模型三者必须一致）===')
  // 批注与模板无关，验一遍即可；但要在第一套模板的版面上验，且要点得到侧栏。
  await selectTemplate(page, first.template)
  // 期望值不从 DOM 反推，而是拿组件的模型 —— 从 DOM 反推只能证明
  // 「DOM 与 DOM 自洽」，证明不了它跟模型一致，而模型才是导出 docx 的依据。
  const commentModel = firstReport.model ?? parseMd(firstReport.source)
  const expectedScopes = commentScopes(commentModel)
  const hasComments = commentModel.comments.length > 0
  eq('侧栏按有无批注出现', firstReport.commentSidebar.present, hasComments)
  eq(
    '侧栏条目数',
    firstReport.commentSidebar.items.length,
    hasComments ? commentModel.comments.length : 0,
  )
  eq('正文批注锚点数', firstReport.commentAnchors.length, commentModel.comments.length)

  for (const c of commentModel.comments) {
    const item = firstReport.commentSidebar.items[c.id]
    const scope = expectedScopes.get(c.id) ?? ''
    const anchor = firstReport.commentAnchors.find((a) => a.id === String(c.id))
    const checks = [
      eq(`批注#${c.id + 1}·内容`, item?.text ?? '', c.text),
      eq(`批注#${c.id + 1}·侧栏锚定文字`, item?.scope ?? '', `「${scope}」`),
      eq(`批注#${c.id + 1}·正文锚点文字`, anchor?.text ?? '', scope),
    ]
    console.log(
      `${checks.every(Boolean) ? 'ok  ' : 'FAIL'} 批注#${c.id + 1} ` +
        `锚定"${anchor?.text ?? ''}" 内容"${item?.text ?? ''}"`,
    )
  }

  // 点侧栏 → 正文锚点高亮。这条是「能看到批注内容」这个需求的落点，必须实测。
  const firstComment = commentModel.comments[0]
  if (firstComment) {
    await page.click('.wtp-comments li .wtp-comment-item')
    // Vue 的更新在微任务里，等一帧再读更稳；超时也让后面的断言给出真实差异
    await page
      .waitForFunction(() => document.querySelectorAll('.wtp-comment-active').length > 0, null, {
        timeout: 3000,
      })
      .catch(() => {})
    const clicked = await page.evaluate(() => ({
      active: document.querySelectorAll('.wtp-comment-active').length,
      text: document.querySelector('.wtp-comment-active')?.textContent ?? '',
    }))
    const expectedAnchor = expectedScopes.get(firstComment.id) ?? ''
    ok('点侧栏条目后正文锚点高亮', clicked.active >= 1, `高亮 ${clicked.active} 处`)
    eq('高亮的是被点的那条锚点', clicked.text, expectedAnchor)
    console.log(`  ok   点第 1 条批注 → 正文高亮"${clicked.text}"`)
  }

  console.log('\n=== 6. 两套模板的差异（切模板必须整篇重量）===')
  ok(
    '两套模板的页数不相同',
    first.report.pageCount !== second.report.pageCount,
    `「${first.template.label}」${first.report.pageCount} 页 vs ` +
      `「${second.template.label}」${second.report.pageCount} 页`,
  )
  ok(
    '公文模板版心更矮 → 页数不少于管理人文件',
    second.report.pageCount >= first.report.pageCount,
    `${first.report.pageCount} → ${second.report.pageCount}`,
  )
  eq(
    '两套模板的内容完全一致（切模板不该动到文档）',
    JSON.stringify(first.report.model?.blocks),
    JSON.stringify(second.report.model?.blocks),
  )
  console.log(
    `  ok   页数：「${first.template.label}」${first.report.pageCount} 页，` +
      `「${second.template.label}」${second.report.pageCount} 页；` +
      `版心宽 ${first.report.pages[0]?.contentWidth} → ${second.report.pages[0]?.contentWidth}px`,
  )

  console.log('\n=== 7. 双页并排（宽视口同排、窄视口回落）===')
  const paperWidth = lengthToPx(first.spec.page.size.width)

  /** 页几何：top/left/width，用来判断「谁和谁在同一行」 */
  const pageGeometry = () =>
    page.evaluate(() => {
      const strip = document.querySelector('.wtp-pages')
      return {
        pages: Array.from(document.querySelectorAll('.wtp-page')).map((el) => {
          const r = el.getBoundingClientRect()
          return { top: Math.round(r.top), left: Math.round(r.left), width: Math.round(r.width) }
        }),
        strip: strip
          ? { clientWidth: Math.round(strip.clientWidth), scrollWidth: Math.round(strip.scrollWidth) }
          : null,
      }
    })

  // 宽视口：A4 两页并排 ≈ 1605px，2200px 视口足够
  await page.setViewportSize({ width: 2200, height: 1100 })
  await selectTemplate(page, first.template)
  await page.waitForTimeout(300)
  const wide = await pageGeometry()
  const [p1, p2] = wide.pages
  ok('宽视口下有两页并排', p1 !== undefined && p2 !== undefined && p1.top === p2.top, JSON.stringify(wide.pages.slice(0, 3)))
  if (p1 && p2) {
    const gap = p2.left - (p1.left + p1.width)
    approx('并排两页的水平间距（flex gap）', gap, 18, 1.5)
    approx('并排第一页仍是整幅纸宽', p1.width, paperWidth, 1)
    approx('并排第二页仍是整幅纸宽', p2.width, paperWidth, 1)
    console.log(
      `  ok   第 1、2 页同一行（top ${p1.top}），间距 ${gap}px，页宽 ${p1.width}/${p2.width}px（纸宽 ${round2(paperWidth)}px）`,
    )
  }
  ok(
    '宽视口下页带没有横向溢出',
    wide.strip !== null && wide.strip.scrollWidth <= wide.strip.clientWidth + 1,
    JSON.stringify(wide.strip),
  )
  // 一份并排的截图，肉眼复核用
  {
    const shot = await page.$('.wtp-root')
    if (shot) {
      const path = join(root, '.qwen', 'tmp', 'side-by-side.png')
      await shot.screenshot({ path })
      console.log(`  ok   并排截图：${path}`)
    }
  }

  // 窄视口：只放得下一页，必须回落成一页一排
  await page.setViewportSize({ width: 1000, height: 1100 })
  await page.waitForTimeout(400)
  const narrow = await pageGeometry()
  const tops = narrow.pages.map((p) => p.top)
  ok('窄视口回落成一页一排（没有两页共用一个 top）', new Set(tops).size === tops.length, JSON.stringify(tops))
  for (const [i, p] of narrow.pages.entries()) {
    approx(`窄视口第${i + 1}页仍是整幅纸宽（没被压窄）`, p.width, paperWidth, 1)
  }
  console.log(
    `  ok   窄视口 ${narrow.pages.length} 页的 top：${tops.join(',')}；页宽 ${narrow.pages[0]?.width}px`,
  )

  await page.setViewportSize({ width: 1700, height: 1100 })
  await page.waitForTimeout(300)

  console.log('\n=== 8. 逐节几何与命名页（横竖混排；W5）===')
  /** 每张纸的几何 + 它挂的命名页 + 页间断页方式 */
  const paperView = () =>
    page.evaluate(() =>
      Array.from(document.querySelectorAll('.wtp-page')).map((el) => {
        const rect = el.getBoundingClientRect()
        return {
          page: el.style.getPropertyValue('page'),
          width: Math.round(rect.width),
          height: Math.round(rect.height),
          contentWidth: el.querySelector('.wtp-content')?.clientWidth ?? 0,
          breakBefore: getComputedStyle(el).breakBefore,
        }
      }),
    )
  /** 把插入符放进含 needle 的那一片段的末尾（verify-browser 里没有 __wtpTest，手工做一遍） */
  const setCaretInText = (needle) =>
    page.evaluate((n) => {
      const frag = Array.from(document.querySelectorAll('.wtp-content > [data-block-id]')).find(
        (el) => (el.textContent ?? '').includes(n),
      )
      if (!frag) return null
      const host = frag.closest('[contenteditable="true"]')
      if (host) host.focus()
      const walker = document.createTreeWalker(frag, NodeFilter.SHOW_TEXT)
      const node = walker.nextNode()
      if (!node) return null
      const range = document.createRange()
      range.setStart(node, node.data.length)
      range.collapse(true)
      const sel = document.getSelection()
      sel.removeAllRanges()
      sel.addRange(range)
      return frag.dataset.blockId ?? null
    }, needle)

  const secBar = page.locator('.section-toolbar')
  ok('编辑模式下「节」工具条常驻', (await secBar.count()) === 1)
  const screenView = await paperView()
  ok(
    '每张纸都挂了命名页（wtp-portrait / wtp-landscape）',
    screenView.length > 0 &&
      screenView.every((p) => p.page === 'wtp-portrait' || p.page === 'wtp-landscape'),
    JSON.stringify(screenView.slice(0, 2)),
  )
  ok(
    '样本默认方向是纵向（宽 < 高）',
    screenView.every((p) => p.width < p.height),
    JSON.stringify(screenView[0]),
  )
  ok(
    '第 2 张纸起都在新的一页开始（一页一张纸靠它）',
    screenView.slice(1).every((p) => p.breakBefore === 'page'),
    JSON.stringify(screenView.map((p) => p.breakBefore)),
  )

  // 打印媒体下：命名页不许把分页改坏（break-before 与页数都得原样）
  await page.emulateMedia({ media: 'print' })
  const printPapers = await paperView()
  await page.emulateMedia({ media: 'screen' })
  eq('打印媒体下纸数不变', printPapers.length, screenView.length)
  ok(
    '打印媒体下命名页仍在每张纸上',
    printPapers.every((p) => p.page === 'wtp-portrait' || p.page === 'wtp-landscape'),
    JSON.stringify(printPapers.map((p) => p.page)),
  )
  ok(
    '打印媒体下第 2 张起仍在新的纸开始（命名页没有引入多余断页）',
    printPapers.slice(1).every((p) => p.breakBefore === 'page'),
    JSON.stringify(printPapers.map((p) => p.breakBefore)),
  )

  // 用「节」工具条把最后一节改成横排：只有那一节的纸横过来，页数不变
  const tailBlock = await setCaretInText('2026年9月12日')
  ok('找得到最后一节里的落点', tailBlock !== null)
  await secBar
    .locator('.tk-group', { hasText: '方向' })
    .getByRole('radio', { name: '横向', exact: true })
    .click()
  await page.waitForTimeout(500)
  const mixed = await paperView()
  const landscape = mixed.filter((p) => p.page === 'wtp-landscape')
  const portrait = mixed.filter((p) => p.page === 'wtp-portrait')
  ok('最后一节变成横排（宽 > 高）', landscape.length > 0 && landscape.every((p) => p.width > p.height), JSON.stringify(mixed))
  ok(
    '其余节仍是纵排（只有那一节被改）',
    portrait.length > 0 && portrait.every((p) => p.width < p.height),
    JSON.stringify(mixed),
  )
  ok(
    '横排纸的宽高恰好是纵排纸的对调',
    landscape[0]?.width === portrait[0]?.height && landscape[0]?.height === portrait[0]?.width,
    `${JSON.stringify(landscape[0])} vs ${JSON.stringify(portrait[0])}`,
  )
  eq('横竖混排纸数不变', mixed.length, screenView.length)
  // 横排纸比页带宽：wrap 之后它独占一行且不许被压窄（压窄 = 版心变窄 = 分页全错）
  ok(
    '横排纸没有被压窄',
    landscape.every((p) => p.width === portrait[0]?.height),
    JSON.stringify(landscape[0]),
  )
  // 还原成纵向，免得影响后面的产出物与截图
  await secBar
    .locator('.tk-group', { hasText: '方向' })
    .getByRole('radio', { name: '纵向', exact: true })
    .click()
  await page.waitForTimeout(400)
  ok(
    '还原后所有纸都是纵排',
    (await paperView()).every((p) => p.page === 'wtp-portrait' && p.width < p.height),
  )

  console.log('\n=== 9. 产出物 ===')
  const tmpDir = join(root, '.qwen', 'tmp')
  mkdirSync(tmpDir, { recursive: true })
  const sourcePath = join(tmpDir, 'demo-source.md')
  const sourceMd = firstReport.model ? toMd(firstReport.model) : firstReport.source
  if (!firstReport.model && !firstReport.source) {
    failures.push('拿不到文档模型也拿不到源码，无法产出 demo-source.md')
  }
  writeFileSync(sourcePath, sourceMd, 'utf8')
  console.log(`  ok   界面源码已导出：${sourcePath}`)

  // 供 verify-page-count.mjs 与 Word 的页数对账：每套模板各一份报告
  const reportPath = join(tmpDir, 'pagination-report.json')
  writeFileSync(
    reportPath,
    JSON.stringify(
      {
        templates: Object.fromEntries(
          seen.map(({ template, report }) => [
            template.key,
            {
              key: template.key,
              label: template.label,
              pageCount: report.pageCount,
              pageNumbers: report.pages.map((p) => p.number),
              contentWidth: report.pages[0]?.contentWidth ?? 0,
              contentHeight: report.pages[0]?.contentHeight ?? 0,
              // 逐节结果（从模型现推，与预览/导出同一个真相源）。
              // 「重排的节数」是 verify:pages 对账页码 1 出现次数的依据 ——
              // 用模型推出来的期望值当 oracle，而不是拿预览自己数自己。
              ...sectionInfo(report, template),
            },
          ]),
        ),
      },
      null,
      2,
    ),
    'utf8',
  )
  console.log(`  ok   分页结果快照（每套模板一份）：${reportPath}`)

  await selectTemplate(page, first.template)
  const shotPath = join(tmpDir, 'preview.png')
  const pagesEl = await page.$('.wtp-root')
  if (pagesEl) {
    await pagesEl.screenshot({ path: shotPath })
    console.log(`  ok   预览截图（含批注侧栏）：${shotPath}`)
  } else {
    failures.push('找不到 .wtp-root，无法截图')
  }
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
  '[PASS] 浏览器实测：两套模板的样式与版心一致、分页无溢出、切模板页数改变、' +
    '宽视口双页并排、窄视口回落一页一排、分节页码已重排、批注侧栏与锚点一致、' +
    '逐页几何按节（改方向后只有那一节的纸横过来且宽高对调、页数不变）、' +
    '每张纸挂命名页且打印媒体下分页不被改坏。',
)
