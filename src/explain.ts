/** `explain`: where a cell's value comes from, as a tree of its precedents and as indented text in the book's locale. */
import type { CellValue } from './contract.js'
import { key, toCellValue, valueAt, type EngineState } from './engine.js'
import { toLocal } from './locale/index.js'
import { say, type Message } from './locale/messages.js'
import { DEFAULT_LOCALE, LOCALE_FILES } from './locale/registry.js'
import { formatAddress, indexToCol, MAX_COLS, MAX_ROWS } from './model.js'
import { ownerOf } from './spill/place.js'
import { quoteSheet } from './syntax.js'
import type { Rng } from './values.js'

/** Reference text for `explain` and reasons: `'Sheet'!B3`, `'Sheet'!D:D`, `'Sheet'!A1:C9`. */
export function rngText(sheetNames: string[], g: Rng): string {
  const sheet = quoteSheet(sheetNames[g.s])
  if (g.r1 === 0 && g.r2 === MAX_ROWS - 1) return `${sheet}!${indexToCol(g.c1)}:${indexToCol(g.c2)}`
  if (g.c1 === 0 && g.c2 === MAX_COLS - 1) return `${sheet}!${g.r1 + 1}:${g.r2 + 1}`
  const a = formatAddress(g.r1, g.c1)
  return g.r1 === g.r2 && g.c1 === g.c2 ? `${sheet}!${a}` : `${sheet}!${a}:${formatAddress(g.r2, g.c2)}`
}

export interface ExplainTree {
  ref: string
  value?: CellValue
  formula?: string
  cells?: number
  /** A cell an array spilled into: the array's formula cell, which is the node's child. */
  spillOf?: string
  children: ExplainTree[]
}

export function explainCell(state: EngineState, s: number, r: number, c: number, depth: number, maxNodes: number, expand: boolean): { root: ExplainTree; nodes: number; truncated: boolean } {
  let nodes = 0
  let truncated = false
  const visit = (s: number, r: number, c: number, d: number): ExplainTree => {
    nodes++
    const id = state.at.get(key(s, r, c))
    const f = id === undefined ? undefined : state.formulas[id]
    const node: ExplainTree = { ref: `${quoteSheet(state.sheetNames[s])}!${formatAddress(r, c)}`, value: toCellValue(valueAt(state, s, r, c)), children: [] }
    const inv = state.invalid.get(key(s, r, c))
    if (inv) node.formula = state.input.sheets[s].cells[formatAddress(r, c)]?.formula
    const owner = f ? undefined : ownerOf(state, s, r, c)
    if (owner) {
      node.spillOf = `${quoteSheet(state.sheetNames[s])}!${formatAddress(owner.r1, owner.c1)}`
      if (d > 0) node.children.push(visit(s, owner.r1, owner.c1, d - 1))
      else truncated = true
    }
    if (!f) return node
    node.formula = f.text
    if (d <= 0) {
      if (f.refs.length) truncated = true
      return node
    }
    for (const g of f.refs) {
      if (nodes >= maxNodes) {
        truncated = true
        break
      }
      if (g.r1 === g.r2 && g.c1 === g.c2) {
        node.children.push(visit(g.s, g.r1, g.c1, d - 1))
        continue
      }
      const b = state.rt.bounds(g)
      const flat = state.rt.flat(g)
      const filled: number[] = []
      flat.forEach((v, i) => v !== null && filled.push(i))
      nodes++
      const rn: ExplainTree = { ref: rngText(state.sheetNames, g), cells: filled.length, children: [] }
      node.children.push(rn)
      if (!expand) continue
      for (const i of filled) {
        if (nodes >= maxNodes) {
          truncated = true
          break
        }
        rn.children.push(visit(g.s, b.r1 + Math.floor(i / b.w), b.c1 + (i % b.w), d - 2))
      }
    }
    return node
  }
  const root = visit(s, r, c, depth)
  return { root, nodes, truncated }
}

/** The tree as indented text in the locale `tag` (en-US without one): its words, values and formulas. */
export function explainText(root: ExplainTree, tag?: string): string {
  const loc = (LOCALE_FILES.get(tag ?? DEFAULT_LOCALE) ?? LOCALE_FILES.get(DEFAULT_LOCALE)!).text
  const local = (f: string) => {
    const t = toLocal(f, loc)
    return t.ok ? t.formula : f
  }
  const text = (m: Message) => say(loc, m, local)
  const lines: string[] = []
  const go = (n: ExplainTree, indent: string) => {
    const head = n.cells !== undefined ? `${n.ref} (${text({ key: 'explainCells', count: n.cells })})` : `${n.ref} = ${text({ key: 'explainValue', value: n.value })}`
    const from = n.formula ? text({ key: 'explainFormula', formula: n.formula }) : n.spillOf ? text({ key: 'explainSpill', ref: n.spillOf }) : ''
    lines.push(`${indent}${head}${from ? `  ← ${from}` : ''}`)
    for (const ch of n.children) go(ch, indent + '  ')
  }
  go(root, '')
  return lines.join('\n')
}
