/**
 * 节的模型操作（纯函数，不碰 DOM，可在 node 里单测）。
 *
 * 分节符只是 blocks 里的一枚分隔符，节的设置另存在 `doc.sections`（下标 = 节号），
 * 两者必须同步 —— 这个同步**只许**走这里的封装函数，别处不要直接 splice blocks 的
 * 分节符，否则会出现「模型里有 2 个分节符、sections 只有 1 项」这类错位。
 */

import { nextBlockId, parseCellId } from '../types'
import type { DocModel, SectionSettings } from '../types'

/**
 * 一处归一化：把「与默认值相等」的字段删掉，并清掉无意义的组合。
 *
 * 三个理由，缺一个都会出问题：
 *   1. 模型只存非默认值 → md 序列化「只写非默认值」才有意义，往返才字节稳定；
 *   2. 首节没有前节，linkPrevious 恒无意义；首节恒从 1 开始，restartAtOne 也无意义；
 *   3. linkPrevious 为真时另两项被它完全接管（互斥），留着就是「说了两句话」。
 *
 * 与表格围栏的 `minLines=1` / `cantSplit=yes` 是同一个约定（见 PLAN 8.9②）。
 */
export function normalizeSectionSettings(raw: SectionSettings, isFirst: boolean): void {
  if (raw.orientation === 'portrait') delete raw.orientation
  if (isFirst) {
    delete raw.linkPrevious
    delete raw.restartAtOne
  }
  const linked = isFirst ? false : (raw.linkPrevious ?? true)
  if (linked) {
    if (!isFirst) delete raw.linkPrevious
    delete raw.pageNumbers
    delete raw.restartAtOne
  } else {
    if (raw.pageNumbers === true) delete raw.pageNumbers
    if (raw.restartAtOne === false) delete raw.restartAtOne
  }
}

/** 全默认的一节（没有任务字段） */
function isEmptySection(s: SectionSettings): boolean {
  return Object.keys(s).length === 0
}

/**
 * 一个文档里有几节 = 分节符数 + 1（空白文档本身就是第一节）。
 * 节号与分节符一一对应：第 k 个分节符开启第 k+1 节。
 */
export function sectionCountOf(doc: DocModel): number {
  let n = 1
  for (const block of doc.blocks) {
    if (block.t === 'sectionBreak') n += 1
  }
  return n
}

/**
 * 光标所在的节号。blockId 是格子 id 时先归到那张表（表格本身不属于任何格）。
 * 找不到（id 过期、块被删）时返回 0：调用方按首节处理，不抛错。
 */
export function sectionIndexOf(doc: DocModel, blockId: string): number {
  const cell = parseCellId(blockId)
  const target = cell ? cell.tableId : blockId
  let index = 0
  for (const block of doc.blocks) {
    if (block.id === target) return index
    if (block.t === 'sectionBreak') index += 1
  }
  return 0
}

/** 某一节的原始设置（可能是 undefined = 全默认） */
export function settingsOf(doc: DocModel, index: number): SectionSettings | undefined {
  return doc.sections?.[index]
}

/**
 * 把 `doc.sections` 补齐到「分节符数 + 1」项（缺项补 `{}`）。
 * 编辑前先调它，后面的 splice 才能对准下标。
 */
function ensureSections(doc: DocModel): SectionSettings[] {
  const count = sectionCountOf(doc)
  const sections = doc.sections ?? (doc.sections = [])
  while (sections.length < count) sections.push({})
  if (sections.length > count) sections.length = count
  return sections
}

/**
 * 全默认时把 `sections` 整个删掉（模型里不留等价于「没写」的空壳）。
 *
 * 这是 md 往返字节稳定的必要条件：`sections: [{}, {}]` 与「没有 sections 字段」
 * 语义完全相同，若两种形态都能存在，`模型 → md → 模型` 就会在这两种形态间漂移。
 */
function pruneSections(doc: DocModel): void {
  const sections = doc.sections
  if (!sections) return
  if (sections.every(isEmptySection)) delete doc.sections
}

/**
 * 改一节的设置。只落非默认值；`undefined` 表示「删掉这个字段」。
 * 首节与「关联前节」组合会被归一化清掉无意义的项（见 normalizeSectionSettings）。
 *
 * 节号越界（调用方算错 / 块被删）时什么都不做：写进去会撑出一个
 * 「sections 比节数还长」的错位数组，把不变式破坏掉。
 */
export function setSectionSetting(
  doc: DocModel,
  index: number,
  patch: SectionSettings,
): void {
  if (index < 0 || index >= sectionCountOf(doc)) return
  const sections = ensureSections(doc)
  const raw: SectionSettings = { ...(sections[index] ?? {}) }
  for (const [key, value] of Object.entries(patch) as [keyof SectionSettings, unknown][]) {
    if (value === undefined) delete raw[key]
    else (raw as Record<string, unknown>)[key] = value
  }
  normalizeSectionSettings(raw, index === 0)
  sections[index] = raw
  pruneSections(doc)
}

/**
 * 在 blockId 之后插入一个分节符，并在 `doc.sections` 的对应位置插入一项（全默认）。
 *
 * **不沿用前一节的设置**（用户 2026-09-13 拍板）：沿用会让「插一个分节符」意外复制
 * 上一节的横向纸张；新节给一个干净的全默认（numbers=on / link=on / restart=off / portrait）
 * 才是可预期的行为，也与 `insertBreakAfter` 原来的表现一致。
 *
 * blockId 找不到（或没给）时追加到文末。
 */
export function insertSectionBreakAfter(doc: DocModel, blockId: string | undefined): string {
  const found = blockId === undefined ? -1 : doc.blocks.findIndex((b) => b.id === blockId)
  const at = found < 0 ? doc.blocks.length : found + 1

  // 新分节符开启的节号 = 它前面已有的分节符数 + 1（块还没插进去，先数一遍）
  let openedIndex = 1
  for (let i = 0; i < at; i += 1) {
    if (doc.blocks[i]?.t === 'sectionBreak') openedIndex += 1
  }

  const sections = ensureSections(doc)
  const block = { t: 'sectionBreak' as const, id: nextBlockId('s') }
  doc.blocks.splice(at, 0, block)
  sections.splice(openedIndex, 0, {})
  // 新节是全默认的；若全文都还是默认，就把 sections 收回去（与 setSectionSetting 同一约定，
  // 否则「插过一个分节符再删掉」会留下一个等价于「没写」的空壳，md 往返就不稳定了）
  pruneSections(doc)
  return block.id
}

/**
 * 删掉一个分节符：同步把**它开启的那一节**的设置项从 `doc.sections` 里摘掉。
 * 不是分节符（段落、分页符、表格）时返回 false 且一个字节都不改。
 */
export function removeSectionBreak(doc: DocModel, blockId: string): boolean {
  const index = doc.blocks.findIndex((b) => b.id === blockId)
  const block = doc.blocks[index]
  if (index < 0 || !block || block.t !== 'sectionBreak') return false

  // 这个分节符开启的节号 = 它前面的分节符数 + 1
  let openedIndex = 1
  for (let i = 0; i < index; i += 1) {
    if (doc.blocks[i]?.t === 'sectionBreak') openedIndex += 1
  }

  const sections = ensureSections(doc)
  doc.blocks.splice(index, 1)
  if (openedIndex < sections.length) sections.splice(openedIndex, 1)
  pruneSections(doc)
  return true
}
