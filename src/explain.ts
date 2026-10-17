/** `explain`: where a cell's value comes from, as a tree of its precedents and as indented text in the book's locale. */
import type { CellValue, ExplainOptions, ExternalBook, RangeRef } from './contract.js'
import { key, toCellValue, valueAt, type EngineState } from './engine.js'
import { object, onlyOptions } from './input.js'
import { errorOrigin, errorStep, issueCode, placeText, type Place } from './issues.js'
import { showFormula } from './locale/index.js'
import { say, type Message } from './locale/messages.js'
import { localeOf } from './locale/registry.js'
import { formatAddress, indexToCol, MAX_COLS, MAX_ROWS, placeOf } from './model.js'
import { runtimeSheet } from './links.js'
import { ownerOf } from './spill/place.js'
import { quoteSheet } from './syntax.js'
import type { Rng } from './values.js'
import { paced, type Steps } from './exec/budget.js'

/** `explain`'s options, checked: an unknown one is refused, as another call's with the call that takes it. */
export function checkExplainOptions(v: unknown, path = 'opts'): ExplainOptions {
  if (v === undefined) return {}
  onlyOptions(object(v, path, 'options such as { depth: 3, maxNodes: 200 }'), path, ['depth', 'maxNodes', 'expand', 'errorDepth'])
  return v as ExplainOptions
}

/** Reference text for `explain` and reasons: `'Sheet'!B3`, `'Sheet'!D:D`, `'Sheet'!A1:C9`. */
export function rngText(state: EngineState, g: Rng): string {
  const sheet = quoteSheet(runtimeSheet(state.sheetNames, state.links, g.s))
  if (g.r1 === 0 && g.r2 === MAX_ROWS - 1) return `${sheet}!${indexToCol(g.c1)}:${indexToCol(g.c2)}`
  if (g.c1 === 0 && g.c2 === MAX_COLS - 1) return `${sheet}!${g.r1 + 1}:${g.r2 + 1}`
  const a = formatAddress(g.r1, g.c1)
  return g.r1 === g.r2 && g.c1 === g.c2 ? `${sheet}!${a}` : `${sheet}!${a}:${formatAddress(g.r2, g.c2)}`
}

export interface ExplainTree {
  location?: RangeRef
  ref: string
  value?: CellValue
  formula?: string
  cells?: number
  /** A cell an array spilled into: the array's formula cell, which is the node's child. */
  spillOf?: string
  /** A cell with an error, or a formula a circular reference stops: the cell it began in. */
  origin?: string
  children: ExplainTree[]
}

/**
 * The tree of a cell's precedents, `depth` levels deep (a range takes one); a cell with an error leads on to the cell it
 * began in, through a range by the cell its error came from, past `depth` by that way alone for `errorDepth` cells more,
 * stopping where it comes back to a cell above it.
 */
export function* explainCell(state: EngineState, s: number, r: number, c: number, o: { depth: number; maxNodes: number; expand: boolean; errorDepth: number }): Steps<{ root: ExplainTree; nodes: number; truncated: boolean }> {
  let nodes = 0
  let truncated = false
  // the cells from the root to the one visited, each as often as it is there; the origins of errors found
  const above = new Map<string, number>()
  const origins = new Map<string, Place>()
  const room = () => nodes < o.maxNodes || !(truncated = true)
  const range = function* (g: Rng): Steps<ExplainTree> {
    nodes++
    let filled = 0
    for (const v of state.rt.flat(g)) {
      if (paced(1)) yield
      if (v !== null) filled++
    }
    const ref = rngText(state, g)
    const location = state.sheetNames[g.s] === undefined ? undefined : { sheet: state.sheetNames[g.s], range: ref.slice(ref.lastIndexOf('!') + 1) }
    return { ref, ...(location ? { location } : {}), cells: filled, children: [] }
  }
  // `d`: the levels left; `past`: the cells the way to an error's origin may go on past them
  const visit = function* (s: number, r: number, c: number, d: number, past: number): Steps<ExplainTree> {
    if (paced(1)) yield
    nodes++
    const k = key(s, r, c)
    const id = state.at.get(placeOf(s, r, c))
    const f = id === undefined ? undefined : state.formulas[id]
    const v = valueAt(state, s, r, c)
    const location = state.sheetNames[s] === undefined ? undefined : { sheet: state.sheetNames[s], range: formatAddress(r, c) }
    const node: ExplainTree = { ref: placeText(state, { s, r, c }), ...(location ? { location } : {}), value: toCellValue(v), children: [] }
    if (state.invalid.has(k)) node.formula = state.input.sheets[s].cells[formatAddress(r, c)]?.formula
    if (f) node.formula = f.text
    const code = issueCode(state, s, r, c)
    const step = code ? errorStep(state, { s, r, c }, code) : null
    if (code) node.origin = placeText(state, errorOrigin(state, { s, r, c }, code, origins))
    // past the depth only the way to the origin goes on, while it does not come back to a cell above
    const onward = step && past > 0 && !above.has(key(step.s, step.r, step.c)) ? step : null
    above.set(k, (above.get(k) ?? 0) + 1)
    const owner = f ? undefined : ownerOf(state, s, r, c)
    if (owner) {
      node.spillOf = placeText(state, { s, r: owner.r1, c: owner.c1 })
      if (d > 0) node.children.push(yield* visit(s, owner.r1, owner.c1, d - 1, past))
      else if (onward && room()) node.children.push(yield* visit(s, owner.r1, owner.c1, 0, past - 1))
      else truncated = true
    }
    if (f && d <= 0) {
      if (onward && room()) {
        const g = f.refs[onward.ref!]
        const single = g.r1 === g.r2 && g.c1 === g.c2
        const into = single ? node : yield* range(g)
        if (!single) node.children.push(into)
        into.children.push(yield* visit(onward.s, onward.r, onward.c, 0, past - 1))
        if (f.refs.length > 1) truncated = true
      } else if (f.refs.length) truncated = true
    } else if (f) {
      for (let i = 0; i < f.refs.length; i++) {
        if (!room()) break
        const g = f.refs[i]
        if (g.r1 === g.r2 && g.c1 === g.c2) {
          node.children.push(yield* visit(g.s, g.r1, g.c1, d - 1, past))
          continue
        }
        const rn = yield* range(g)
        node.children.push(rn)
        if (o.expand) {
          const b = state.rt.bounds(g)
          const flat = state.rt.flat(g)
          for (let j = 0; j < flat.length; j++) {
            if (flat[j] === null) continue
            if (!room()) break
            rn.children.push(yield* visit(g.s, b.r1 + Math.floor(j / b.w), b.c1 + (j % b.w), d - 2, past))
          }
        } else if (step?.ref === i && room()) rn.children.push(yield* visit(step.s, step.r, step.c, d - 2, past))
      }
    }
    const n = above.get(k)!
    if (n > 1) above.set(k, n - 1)
    else above.delete(k)
    return node
  }
  const root = yield* visit(s, r, c, o.depth, o.errorDepth)
  return { root, nodes, truncated }
}

/** The tree as indented text in the loaded locale `tag` (en-US without one): its words, values, formulas and references, a sheet in quotes where Excel puts them. */
export function* explainText(root: ExplainTree, tag?: string, externalBooks?: ExternalBook[]): Steps<string> {
  const loc = localeOf(tag)
  const local = (f: string) => {
    const t = showFormula(f, loc, undefined, externalBooks)
    return t.ok ? t.formula : f
  }
  // the formula shown before its message takes it: en-US's prints it as given
  const text = (m: Message) => say(loc, m)
  const lines: string[] = []
  const go = function* (n: ExplainTree, indent: string): Steps<void> {
    if (paced(1)) yield
    const ref = local(n.ref)
    const head = n.cells !== undefined ? `${ref} (${text({ key: 'explainCells', count: n.cells })})` : `${ref} = ${text({ key: 'explainValue', value: n.value })}`
    const from = n.formula ? text({ key: 'explainFormula', formula: local(n.formula) }) : n.spillOf ? text({ key: 'explainSpill', ref: local(n.spillOf) }) : ''
    // a node with an origin and a value other than an error is a formula a circular reference stops, which holds its value
    const held = n.value === null || typeof n.value !== 'object'
    const origin = n.origin === undefined ? '' : n.origin === n.ref ? text({ key: held ? 'explainHeldHere' : 'explainErrorHere' }) : text(held ? { key: 'explainHeldFrom', ref: local(n.origin) } : { key: 'explainErrorFrom', ref: local(n.origin) })
    lines.push(`${indent}${head}${from ? `  ← ${from}` : ''}${origin ? `  ${origin}` : ''}`)
    for (const ch of n.children) yield* go(ch, indent + '  ')
  }
  yield* go(root, '')
  return lines.join('\n')
}
