/**
 * 标题自动编号。
 *
 * 编号按需求是「写成正文文字」而不是 Word 原生多级列表：
 * 预览和 docx 里都能看到真实编号，重新导入我们自己生成的 docx 时也能准确还原层级。
 * 代价是在 Word 里手动增删标题不会自动重编号，编号只在本组件内维护。
 */

import type { Block } from './types'
import type { NumberingStyle } from './spec'

const CN_DIGITS = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九'] as const

/** 阿拉伯数字转中文序数，支持到 999。超出范围时原样返回数字串。 */
export function chineseNum(n: number): string {
  if (!Number.isInteger(n) || n <= 0) return String(n)
  if (n < 10) return CN_DIGITS[n] ?? String(n)
  if (n < 20) return n === 10 ? '十' : `十${CN_DIGITS[n - 10] ?? ''}`
  if (n < 100) {
    const tens = Math.floor(n / 10)
    const ones = n % 10
    return `${CN_DIGITS[tens] ?? ''}十${ones === 0 ? '' : (CN_DIGITS[ones] ?? '')}`
  }
  if (n < 1000) {
    const hundreds = Math.floor(n / 100)
    const rest = n % 100
    const head = `${CN_DIGITS[hundreds] ?? ''}百`
    if (rest === 0) return head
    if (rest < 10) return `${head}零${CN_DIGITS[rest] ?? ''}`
    return head + chineseNum(rest)
  }
  return String(n)
}

/** 把序号套进编号样式，得到最终写在段首的前缀 */
export function numberingPrefix(style: NumberingStyle, n: number): string {
  switch (style) {
    case 'chineseDot':
      return `${chineseNum(n)}、`
    case 'parenChinese':
      return `（${chineseNum(n)}）`
    case 'arabicDot':
      return `${n}、`
    case 'none':
      return ''
  }
}

/**
 * 遍历全篇，算出每个标题块应该显示的编号前缀。
 * h1 递增时把 h2/h3 归零，h2 递增时把 h3 归零 —— 与需求里「一、／（一）／1、」的层级一致。
 */
export function computeNumbering(
  blocks: readonly Block[],
  styleOf: (block: Block) => NumberingStyle,
): Map<string, string> {
  const result = new Map<string, string>()
  let c1 = 0
  let c2 = 0
  let c3 = 0

  for (const block of blocks) {
    if (block.t !== 'textBlock') continue
    const style = styleOf(block)
    if (block.kind === 'h1') {
      c1 += 1
      c2 = 0
      c3 = 0
      result.set(block.id, numberingPrefix(style, c1))
    } else if (block.kind === 'h2') {
      c2 += 1
      c3 = 0
      result.set(block.id, numberingPrefix(style, c2))
    } else if (block.kind === 'h3') {
      c3 += 1
      result.set(block.id, numberingPrefix(style, c3))
    }
  }

  return result
}

const H1_NUM_RE = /^\s*(?:[一二三四五六七八九十百零]+、)\s*/
const H2_NUM_RE = /^\s*（[一二三四五六七八九十百零]+）\s*/
const H3_NUM_RE = /^\s*\d+、\s*/

/**
 * 导入 docx 时剥掉段首已有的编号，避免重新编号后出现「一、一、」。
 * 只在编号确实是自动编号体系（即该级样式带 numbering）时才剥。
 */
export function stripAutoNumber(kind: string, text: string): string {
  if (kind === 'h1') return text.replace(H1_NUM_RE, '')
  if (kind === 'h2') return text.replace(H2_NUM_RE, '')
  if (kind === 'h3') return text.replace(H3_NUM_RE, '')
  return text
}
