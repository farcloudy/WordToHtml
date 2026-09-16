/**
 * 可配置的快捷键表。
 *
 * 全是与界面无关的纯函数（能在 node 里单测）：把「动作名 → 组合键字符串」解析成结构化的
 * 组合键，再拿键盘事件去比对。动作集就是编辑器 onKeydown 里那批可配置的键 ——
 * **Tab 与方向键不在内**：它们是编辑器手感（行 / 列跨格移动），改了会和浏览器行为打架。
 *
 * 组合键的写法刻意宽容：
 *   · 大小写不敏感（`ctrl+b` / `Ctrl+B` 同一个）；
 *   · `Cmd` / `Command` / `Meta` / `⌘` / `Win` / `Super` 一律归一到「mod」（= `ctrlKey || metaKey`）；
 *   · 数字键比对时同时看 `event.key` 与 `event.code`（沿用 alt+4 那处的写法）——
 *     中文输入法或非美式布局下，按数字键的 `event.key` 可能不是数字字符。
 */

import defaultsFile from './shortcuts.json'

/** 动作名。**顺序即冲突时的优先序**（先到先得），所以不要随手调整 */
export const SHORTCUT_ACTIONS = [
  'bold',
  'underline',
  'trackChanges',
  'find',
  'replace',
  'undo',
  'redo',
  'formatAmount',
  'repeat',
] as const

export type ShortcutAction = (typeof SHORTCUT_ACTIONS)[number]

/**
 * 默认表。**唯一真相源是旁边的 `shortcuts.json`** —— 那是「动作名 → 组合键字符串」的纯映射，
 * 可以直接手改（`npm run dev` 里改完会自己重载）；使用方想换一套键，把那个文件复制过去改好、
 * 当 `shortcuts` prop 传进来即可 —— 下面这个常量**本身就是那个 prop 的默认值**。
 *
 * 这里刻意逐个键写出、而不是整表展开：
 *   · json 里**拼错动作名**（写成 `bolld`）时，`vue-tsc` 会直接报「该属性不存在」——
 *     不会让一个错键悄悄变成「少绑一个动作」；
 *   · 漏了动作同理（`Record<ShortcutAction, string>` 会要求补齐）。
 */
export const DEFAULT_SHORTCUTS: Record<ShortcutAction, string> = {
  bold: defaultsFile.bold,
  underline: defaultsFile.underline,
  trackChanges: defaultsFile.trackChanges,
  find: defaultsFile.find,
  replace: defaultsFile.replace,
  undo: defaultsFile.undo,
  redo: defaultsFile.redo,
  formatAmount: defaultsFile.formatAmount,
  repeat: defaultsFile.repeat,
}

/*
 * 手改 json 的两道护栏，模块加载时各跑一次（都在控制台留一句话，不抛异常、不阻断启动）：
 *   · json 里有认不出的动作名 → 报出来（上面那种写法只挡得住「拼错」这一半，
 *     多写一个 `_comment` 之类的键另一半就漏过去了）；
 *   · 组合键字符串认不出来 → 报出来，否则那个动作会静默变成「按什么都没反应」。
 * 两道都不改默认表本身：坏值就是坏值，报出来比悄悄兜底好查。
 *
 * ⚠️ 位置有讲究：**必须落在 `parseCombo` 与它那批 `const`（KEY_RE / MOD_NAMES…）之后**。
 * 放到文件开头会踩 const 的暂时性死区 —— 打包后首屏 import 就抛
 * `Cannot read properties of undefined (reading 'test')`（这条实测踩过）。
 */
function warnAboutDefaultsFile(): void {
  for (const name of Object.keys(defaultsFile)) {
    if (!(SHORTCUT_ACTIONS as readonly string[]).includes(name)) {
      console.warn(`[WordToHtml] shortcuts.json 里有未知动作名「${name}」，已忽略`)
    }
  }
  for (const action of SHORTCUT_ACTIONS) {
    if (!parseCombo(DEFAULT_SHORTCUTS[action])) {
      console.warn(
        `[WordToHtml] shortcuts.json 里「${action}」的组合键「${DEFAULT_SHORTCUTS[action]}」认不出来，该动作不会绑定任何键`,
      )
    }
  }
}

export interface Combo {
  /** ctrl 或 meta（Cmd / Win / ⌘）—— 两者不区分 */
  mod: boolean
  shift: boolean
  alt: boolean
  /** 主键：小写。单字符（'b' / '4'）或功能键（'f4'） */
  key: string
}

/** 动作 → 组合键。`null` = 这一项没有绑定（冲突的输家，或写坏了） */
export type ShortcutTable = Record<ShortcutAction, Combo | null>

/** 调用方递进来的那张表：只写要改的动作，其余留空即用默认 */
export type ShortcutOverrides = Partial<Record<ShortcutAction, string>>

const MOD_NAMES = new Set(['ctrl', 'control', 'cmd', 'command', 'meta', '⌘', 'win', 'super'])
const SHIFT_NAMES = new Set(['shift', '⇧'])
const ALT_NAMES = new Set(['alt', 'option', '⌥'])
/** 允许的主键：单个字母/数字，或 F1–F24。其余（多字符、标点串）一律不认，宁可报出来 */
const KEY_RE = /^(?:[a-z0-9]|f(?:[1-9]|1[0-9]|2[0-4]))$/

/**
 * 组合键字符串 → 结构。认不出来返回 `null`（调用方据此保留默认值 / 提示）。
 * 拆不动、主键不合法、出现认不出的修饰键，都算「认不出来」——
 * 不猜、不忽略，否则 `Foo+B` 会被悄悄当成加粗。
 */
export function parseCombo(text: string): Combo | null {
  const parts = text
    .split('+')
    .map((part) => part.trim().toLowerCase())
    .filter((part) => part !== '')
  const key = parts[parts.length - 1]
  if (parts.length === 0 || key === undefined || !KEY_RE.test(key)) return null
  const combo: Combo = { mod: false, shift: false, alt: false, key }
  for (const part of parts.slice(0, -1)) {
    if (MOD_NAMES.has(part)) combo.mod = true
    else if (SHIFT_NAMES.has(part)) combo.shift = true
    else if (ALT_NAMES.has(part)) combo.alt = true
    else return null
  }
  return combo
}

/** 两道护栏在这里落地 —— 位置的理由见函数自己的注释 */
warnAboutDefaultsFile()

/** 组合键的规范签名（修饰键顺序固定）。既用于判重，也用于提示文案 */
export function comboLabel(combo: Combo): string {
  const parts: string[] = []
  if (combo.mod) parts.push('Ctrl')
  if (combo.shift) parts.push('Shift')
  if (combo.alt) parts.push('Alt')
  parts.push(combo.key.toUpperCase())
  return parts.join('+')
}

/** 键盘事件里比对用得上的那几个字段（KeyboardEvent 天然满足） */
export interface ShortcutEvent {
  key: string
  code: string
  ctrlKey: boolean
  metaKey: boolean
  shiftKey: boolean
  altKey: boolean
}

/** 这次按键是不是命中了组合键。修饰键**逐项严格比对**（多按了 Shift 就不算命中） */
export function matchShortcut(event: ShortcutEvent, combo: Combo | null | undefined): boolean {
  if (!combo) return false
  if ((event.ctrlKey || event.metaKey) !== combo.mod) return false
  if (event.shiftKey !== combo.shift) return false
  if (event.altKey !== combo.alt) return false
  if (event.key.toLowerCase() === combo.key) return true
  // 数字键再看一眼 event.code：布局 / 输入法可能让 event.key 不是数字字符
  return /^[0-9]$/.test(combo.key) && event.code === `Digit${combo.key}`
}

/**
 * 默认表 + 调用方的覆盖 → 实际生效的表（九个动作都有值，没绑上的为 `null`）。
 *
 * 三条规则（都有单测）：
 *   · **未知动作名忽略 + 一句 warn**，绝不悄悄改默认表 —— 用户写了 `{ bolld: 'Ctrl+B' }`，
 *     ctrl+B 仍要加粗；
 *   · 组合键字符串认不出来：warn 一句、该动作保留默认值；
 *   · 两个动作绑同一个组合键：按 `SHORTCUT_ACTIONS` 的固定顺序**先到先得**，后到的解绑 + warn。
 */
export function resolveShortcuts(override?: ShortcutOverrides): ShortcutTable {
  const table = {} as ShortcutTable
  for (const action of SHORTCUT_ACTIONS) table[action] = parseCombo(DEFAULT_SHORTCUTS[action])

  for (const action of SHORTCUT_ACTIONS) {
    const text = override?.[action]
    if (text === undefined) continue
    const combo = parseCombo(text)
    if (!combo) {
      console.warn(
        `[WordToHtml] 快捷键「${action}」的组合键「${text}」认不出来，这一项仍用默认值`,
      )
      continue
    }
    table[action] = combo
  }

  if (override) {
    for (const name of Object.keys(override)) {
      if (!(SHORTCUT_ACTIONS as readonly string[]).includes(name)) {
        console.warn(`[WordToHtml] 快捷键表里有未知动作名「${name}」，已忽略（默认表不受影响）`)
      }
    }
  }

  const claimed = new Map<string, ShortcutAction>()
  for (const action of SHORTCUT_ACTIONS) {
    const combo = table[action]
    if (!combo) continue
    const label = comboLabel(combo)
    const owner = claimed.get(label)
    if (owner) {
      console.warn(
        `[WordToHtml] 「${owner}」与「${action}」都绑在 ${label} 上，按动作顺序先到先得，「${action}」已解绑`,
      )
      table[action] = null
      continue
    }
    claimed.set(label, action)
  }
  return table
}
