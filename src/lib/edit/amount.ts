/**
 * 金额格式化（纯函数，可在 node 里单测）。
 *
 * 把选中的文字解析成一个数，再写成「千分位 + 固定两位小数」这一种形态，
 * 例如 12345.6 → 12,345.60、1234 → 1,234.00。已经是这个形态的（12,345.60）
 * 再解析一次结果不变，这样反复按快捷键不会越改越乱。
 *
 * 刻意不用 toLocaleString：它按运行环境的 locale 选分组符与小数点，同一段文字
 * 在不同机器上会输出成 12,345.60 或 12.345,60。这里手工拼字符串，结果只取决于
 * 数字本身 —— 公文里的金额不该随打开它的机器变。
 *
 * 不用 Number 做解析与进位：二进制浮点会把 8.615 这类值算成 8.614999…，
 * 直接按十进制字符处理才没有这个坑（金额的第四位小数要能正确进位）。
 */

/** 允许：负号、千分位逗号、至多一个小数点。不接受科学计数法、全角数字、空格 */
const AMOUNT_RE = /^(-?)(\d+(?:,\d{3})*)(?:\.(\d+))?$/

const ZERO = '0'.charCodeAt(0)

/** 把一串十进制数字加 1（用于四舍五入的进位） */
function increment(digits: string): string {
  const out = digits.split('')
  for (let i = out.length - 1; i >= 0; i -= 1) {
    const d = (out[i] as string).charCodeAt(0) - ZERO
    if (d < 9) {
      out[i] = String(d + 1)
      return out.join('')
    }
    out[i] = '0'
  }
  return `1${out.join('')}`
}

/** 整数部分按三位一组插逗号 */
function group(digits: string): string {
  let out = ''
  for (let i = 0; i < digits.length; i += 1) {
    if (i > 0 && (digits.length - i) % 3 === 0) out += ','
    out += digits[i]
  }
  return out
}

/**
 * 解析并格式化。返回 null 表示「不是合法数字」，调用方据此提示无效。
 * 合法输入一定返回非空字符串（包括负号、零）。
 */
export function formatAmount(raw: string): string | null {
  const text = raw.trim()
  if (text === '') return null
  const m = AMOUNT_RE.exec(text)
  if (!m) return null

  const sign = m[1] === '-' ? '-' : ''
  // 去掉前导零（0001234 → 1234），否则分组会写成 0,001,234
  const int = (m[2] ?? '').replace(/,/g, '').replace(/^0+(?=\d)/, '')
  const frac = m[3] ?? ''
  if (int === '') return null

  const kept = frac.slice(0, 2).padEnd(2, '0')
  // 第三位小数 ≥ 5 就进位（金额只到分）
  const third = frac.length > 2 ? frac.charCodeAt(2) - ZERO : 0
  const rounded = third >= 5 ? increment(int + kept) : int + kept

  // 进位可能让整数部分多一位（999.999 → 1000.00），所以进位后再拆
  const cents = rounded.slice(-2)
  const yuan = rounded.slice(0, -2) || '0'

  // 四舍五入后归零的负数（-0.001）写成 0.00，不写成 -0.00
  const negative = sign === '-' && !/^0*$/.test(yuan + cents)
  return `${negative ? '-' : ''}${group(yuan)}.${cents}`
}
