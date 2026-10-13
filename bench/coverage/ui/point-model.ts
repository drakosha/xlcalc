/**
 * The sheet component's model of Point mode, as its predictions of the set `point-held` (`xlcalc-sheet` master of
 * 14.10.2026: `src/core/point.ts` `refFits` and `usePoint().key`, `src/core/editing.ts` `onEditorKey`, `inLiteral` of
 * `src/core/autocomplete.ts`), ported here. In a field typing opened (Enter mode) an arrow where a reference may stand
 * puts the address of the cell beside the edited one, or moves the reference it put from its first cell, Shift its far
 * corner; an arrow anywhere else in a formula moves the caret, in text commits and moves. F2 turns Enter mode to Edit
 * mode and back; in Edit mode arrows move the caret. Command and an arrow, Home and Page Down are the field's: Command
 * and Right or Down take the caret to the end, Left or Up to the start; Home and Page Down scroll (Mac). Return commits
 * and moves down. The list of functions, which takes arrows while it is open, is not modelled.
 */
import { createSession, rangeAddress, type BookInput } from '../../../src/index.js'
import { formatAddress, indexToCol, parseAddress } from '../../../src/model.js'
import type { PointStep } from './cases.js'

const BEFORE_REF = new Set([...'=+-*/^&<>(:@'])

/** Inside a string, a quoted sheet's name or a table's brackets, up to `upTo`. */
function inLiteral(text: string, upTo: number): boolean {
  let quote: string | null = null
  let depth = 0
  for (let i = 1; i < upTo; i++) {
    const c = text[i]
    if (quote) {
      if (c === quote) quote = null
    } else if (c === '"' || c === "'") quote = c
    else if (c === '[') depth++
    else if (c === ']') depth = Math.max(0, depth - 1)
  }
  return quote !== null || depth > 0
}

/** Whether a reference may stand at `at`: after `BEFORE_REF` or the list separator, spaces between, not in a literal. */
function refFits(text: string, at: number, separator: string): boolean {
  if (!text.startsWith('=') || at < 1 || at > text.length) return false
  let i = at
  while (i > 1 && /\s/.test(text[i - 1])) i--
  const c = text[i - 1]
  return (BEFORE_REF.has(c) || c === separator) && !inLiteral(text, i - 1)
}

/** A cell's place from 0, as `parseAddress` gives it. */
interface Pos {
  row: number
  col: number
}
const STEPS: Record<string, [number, number]> = { down: [1, 0], up: [-1, 0], right: [0, 1], left: [0, -1] }
const moved = (p: Pos, [dr, dc]: [number, number]): Pos => ({ row: Math.max(0, p.row + dr), col: Math.max(0, p.col + dc) })

/** The formula bar after each step, as the harness reads its description (`Formula Bar. C,3 =D3`), and the book after. */
export function pointPrediction(book: BookInput, at: string, steps: PointStep[], tag: string, cells: string[]) {
  const s = createSession(book, { clock: { now: '2026-10-14T12:00:00Z', seed: 1 } })
  try {
    const separator = tag === 'ru-RU' ? ';' : ','
    const a = parseAddress(at)
    let active: Pos = { row: a.row, col: a.col }
    let ed: { text: string; caret: number; mode: 'enter' | 'edit'; at: Pos } | null = null
    // set inside the closures below, so typed by a cast, not narrowed to null
    let run = null as { from: number; to: number; text: string; anchor: Pos; far: Pos } | null
    const shown = (p: Pos) => {
      const addr = formatAddress(p.row, p.col)
      return s.getViewport('S', `A1:${addr}`, { editText: true }).cells[addr]?.editText ?? ''
    }
    const bar = () => {
      const p = ed ? ed.at : active
      return `Formula Bar. ${indexToCol(p.col)},${p.row + 1} ${ed ? ed.text : shown(p)}`
    }
    const commit = (to: [number, number]) => {
      s.setCellText('S', formatAddress(ed!.at.row, ed!.at.col), ed!.text)
      active = moved(ed!.at, to)
      ed = null
      run = null
    }
    const write = (anchor: Pos, far: Pos) => {
      const r = run ?? { from: ed!.caret, to: ed!.caret, text: ed!.text, anchor, far }
      const ref = rangeAddress({ top: Math.min(anchor.row, far.row) + 1, left: Math.min(anchor.col, far.col) + 1, bottom: Math.max(anchor.row, far.row) + 1, right: Math.max(anchor.col, far.col) + 1 })
      ed!.text = r.text.slice(0, r.from) + ref + r.text.slice(r.to)
      ed!.caret = r.from + ref.length
      run = { from: r.from, to: ed!.caret, text: ed!.text, anchor, far }
    }
    const out: { bar: string }[] = []
    for (const step of steps) {
      if ('type' in step) {
        if (!ed) ed = { text: '', caret: 0, mode: 'enter', at: { ...active } }
        ed.text = ed.text.slice(0, ed.caret) + step.type + ed.text.slice(ed.caret)
        ed.caret += step.type.length
        run = null
      } else if (!ed) {
        // a key on the grid: arrows move the selection, Return moves down
        if (STEPS[step.key]) active = moved(active, STEPS[step.key])
        else if (step.key === 'return') active = moved(active, [1, 0])
      } else if (step.key === 'return') commit([1, 0])
      else if (step.key === 'escape') {
        active = ed.at
        ed = null
        run = null
      } else if (step.key === 'f2') ed.mode = ed.mode === 'enter' ? 'edit' : 'enter'
      else if (STEPS[step.key] && !step.cmd && ed.mode === 'enter') {
        const going = run && run.text === ed.text && ed.caret === run.to ? run : null
        if (going || refFits(ed.text, ed.caret, separator)) {
          const extend = going && step.shift
          const anchor = extend ? going.anchor : moved(going ? going.anchor : ed.at, STEPS[step.key])
          write(anchor, extend ? moved(going.far, STEPS[step.key]) : anchor)
        } else if (!ed.text.startsWith('=')) commit(STEPS[step.key])
        else if (step.key === 'left') ed.caret = Math.max(0, ed.caret - 1)
        else if (step.key === 'right') ed.caret = Math.min(ed.text.length, ed.caret + 1)
      } else if (STEPS[step.key]) {
        // the field's own: Command takes the caret to an end, a plain arrow in Edit mode moves it by one
        if (step.cmd) ed.caret = step.key === 'right' || step.key === 'down' ? ed.text.length : 0
        else if (step.key === 'left') ed.caret = Math.max(0, ed.caret - 1)
        else if (step.key === 'right') ed.caret = Math.min(ed.text.length, ed.caret + 1)
      }
      out.push({ bar: bar() })
    }
    const after = Object.fromEntries(cells.map((x) => [x, s.getBook().sheets[0].cells[x] ?? null]))
    return { steps: out, after }
  } finally {
    s.close()
  }
}
