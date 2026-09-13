/**
 * 查找与替换的纯函数层（不碰 DOM，可在 node 里单测）。
 *
 * 坐标系与 edit/model.ts 一致：偏移是**模型文字坐标**，不含标题自动编号前缀。
 * 查找跑在每个块的「可搜索文字」上 —— 也就是把块内文字按顺序连起来、但跳过删除修订
 * 的文字。批注锚点不占字符，格式切换也不占字符，所以它们都不会让匹配漏掉。
 *
 * 默认大小写敏感（与公文场景一致，不做「忽略大小写」选项）。
 */

import type { DocModel, RevMark, TextBlock, TextInline } from '../types'
import { blockLength, deleteRange, findBlock, replaceRange } from './model'

/** 匹配区间（模型文字坐标） */
export interface Match {
  blockId: string
  from: number
  to: number
}

/** 搜索范围（模型文字坐标）。匹配必须完整落在某一个区间内，不许跨区间。 */
export interface SearchScope {
  blockId: string
  from: number
  to: number
}

export interface SearchOptions {
  regex?: boolean
  /** 不给 = 全文；给了空数组 = 没有任何区间，匹配不到任何东西 */
  scope?: SearchScope[]
}

/** 正则报错文案太长会把提示条撑爆，截断到这个长度 */
const MAX_ERROR_LENGTH = 120

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

/** 非法正则返回错误文案（给提示条用），合法返回 null。query 为空按合法处理。 */
export function validateQuery(query: string, regex: boolean): string | null {
  if (!regex || query === '') return null
  try {
    new RegExp(query, 'g')
    return null
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err)
    return message.length > MAX_ERROR_LENGTH ? `${message.slice(0, MAX_ERROR_LENGTH)}…` : message
  }
}

/**
 * 一段连续的可搜索文字（在模型坐标里的起点 + 文字）。
 * 删除修订的文字会把前后断开成两段：查找不该跨过它，替换更不该把它一起吃掉。
 */
interface SearchRun {
  start: number
  text: string
}

function searchRuns(block: TextBlock): SearchRun[] {
  const runs: SearchRun[] = []
  let cursor = 0
  let current: SearchRun | null = null

  for (const inline of block.inlines) {
    if (inline.t !== 'text') continue
    const start = cursor
    cursor += inline.text.length
    if (inline.rev?.kind === 'del') {
      // 修订模式下「替换」是按 Word 语义把旧文字标成 w:del 再插入新文字，旧文字仍留在
      // 模型里；若查找还命中它，用户对同一处反复替换就会不断叠加插入。直接跳过才收敛。
      current = null
      continue
    }
    if (current !== null && current.start + current.text.length === start) {
      current.text += inline.text
    } else {
      current = { start, text: inline.text }
      runs.push(current)
    }
  }

  return runs
}

function collectInRun(blockId: string, run: SearchRun, re: RegExp): Match[] {
  const out: Match[] = []
  re.lastIndex = 0
  let hit = re.exec(run.text)
  while (hit !== null) {
    const matched = hit[0] ?? ''
    if (matched !== '') {
      out.push({ blockId, from: run.start + hit.index, to: run.start + hit.index + matched.length })
    }
    // 零宽匹配没有任何可高亮、可替换的文字，不进结果；但游标必须前进一位，
    // 否则 `^`、`a*` 这类模式会在同一个位置上转不出来。
    re.lastIndex = matched === '' ? hit.index + 1 : hit.index + matched.length
    hit = re.exec(run.text)
  }
  return out
}

/** 查询为空、正则可编译但匹配不到、非法正则 —— 一律返回 []，不抛异常。 */
export function findMatches(doc: DocModel, query: string, opts: SearchOptions = {}): Match[] {
  if (query === '') return []
  const regex = opts.regex === true
  if (regex && validateQuery(query, true) !== null) return []
  const re = new RegExp(regex ? query : escapeRegExp(query), 'g')

  const scopes = new Map<string, SearchScope[]>()
  if (opts.scope) {
    for (const scope of opts.scope) {
      const list = scopes.get(scope.blockId)
      if (list) list.push(scope)
      else scopes.set(scope.blockId, [scope])
    }
  }
  const limited = opts.scope !== undefined

  const out: Match[] = []
  for (const block of doc.blocks) {
    if (block.t !== 'textBlock') continue
    const allowed = limited ? scopes.get(block.id) : undefined
    if (limited && (!allowed || allowed.length === 0)) continue
    for (const run of searchRuns(block)) {
      for (const match of collectInRun(block.id, run, re)) {
        // 区间限定：整段匹配都落在同一个 scope 里才算命中
        if (allowed && !allowed.some((s) => match.from >= s.from && match.to <= s.to)) continue
        out.push(match)
      }
    }
  }
  return out
}

/** 替换结果继承被替换区间内第一个 text 片段的格式：换掉一段加粗文字不该静默丢格式 */
function firstTextFormat(
  block: TextBlock,
  from: number,
  to: number,
): { bold?: boolean; underline?: boolean; color?: string } {
  const fmt: { bold?: boolean; underline?: boolean; color?: string } = {}
  let cursor = 0
  for (const inline of block.inlines) {
    if (inline.t !== 'text') continue
    const start = cursor
    const end = cursor + inline.text.length
    cursor = end
    if (end <= from || start >= to) continue
    if (inline.bold) fmt.bold = true
    if (inline.underline) fmt.underline = true
    if (inline.color) fmt.color = inline.color
    return fmt
  }
  return fmt
}

/**
 * 按 matches 逐处替换成 text，返回替换处数。
 *
 * 块内从后往前处理：前面的替换会改变后面区间的坐标，倒着来才不用维护偏移。
 * 替换文字按纯文本处理，不做反向引用替换（不支持 `$1`）。
 */
export function replaceMatches(
  doc: DocModel,
  matches: readonly Match[],
  text: string,
  makeRev?: (m: Match) => RevMark,
): number {
  const byBlock = new Map<string, Match[]>()
  for (const match of matches) {
    const list = byBlock.get(match.blockId)
    if (list) list.push(match)
    else byBlock.set(match.blockId, [match])
  }

  let count = 0
  for (const [blockId, list] of byBlock) {
    const block = findBlock(doc, blockId)
    if (!block) continue
    list.sort((a, b) => b.from - a.from)
    for (const match of list) {
      const len = blockLength(block)
      const from = Math.max(0, Math.min(match.from, len))
      const to = Math.max(from, Math.min(match.to, len))
      if (to <= from) continue
      const piece: TextInline = { t: 'text', text, ...firstTextFormat(block, from, to) }
      if (makeRev) {
        // kind 由这里按用途覆盖，调用方只需给出唯一 id / 作者 / 时间
        const delMark: RevMark = { ...makeRev(match), kind: 'del' }
        const insMark: RevMark = { ...makeRev(match), kind: 'ins' }
        deleteRange(doc, blockId, from, to, delMark)
        // 新文字落在被删文字之后：读起来是「删除线 + 新文字」，与 Word 一致
        replaceRange(block, to, to, [{ ...piece, rev: insMark }])
      } else {
        replaceRange(block, from, to, [piece])
      }
      count += 1
    }
  }
  return count
}
