/**
 * Excel's Find in a session (`Session.findCells`), measured in Excel for Windows (`bench/coverage/find/`): the query read
 * as it is or with Find's wildcards, matched against a part or the whole of what a cell shows (Look in Values), the text
 * it opens with for editing (Look in Formulas), its stored formula or its value, in any case or in its own; the cells of
 * each range by rows or by columns, those of hidden rows and columns skipped as Find skips them.
 */
import type { CellValue, FindCellsOptions } from '../contract.js'
import { wildSource } from '../criteria.js'
import { key, toCellValue, valueAt, type EngineState } from '../engine.js'
import { checkBoolean, InputError } from '../input.js'
import { isHiddenRow } from '../layout/hidden.js'
import { sheetLines, type SheetLines } from '../layout/lines.js'
import { lines } from '../layout/spans.js'
import { formatAddress, MAX_COLS, MAX_ROWS } from '../model.js'
import { outValue } from '../result.js'
import { cellLooks } from '../shown.js'
import { appearanceOf, windowLooks } from '../cond/looks.js'
import { ownerOf } from '../spill/place.js'
import { hasCell, type Rect } from '../structure/area.js'
import { paced } from '../exec/budget.js'
import { lowerBound } from '../sorted.js'
import { editTexts } from './viewport.js'

/** `findCells`' own options, beside the page's and `fields`. */
export const FIND_OPTIONS = ['in', 'caseSensitive', 'wildcards', 'wholeCell', 'byColumns', 'skipHidden'] as const

const WHERE = ['text', 'edit', 'formula', 'value'] as const

type At = [s: number, r: number, c: number]
type Ref = Rect & { s: number }

/** Each character's lower case as one UTF-16 unit or two as it was, so `?` keeps its place: `İ` stays (its lower case is two characters), `Σ` is `σ` anywhere. */
const folded = (t: string) => (/[İΣ]/.test(t) ? t.replace(/[^İΣ]+/g, (m) => m.toLowerCase()).replaceAll('Σ', 'σ') : t.toLowerCase())

/**
 * The test of a cell's text, null when nothing can match (an empty query, or `~` alone). Wildcards as the lookup
 * functions read them (`wildSource`) but `?` is one UTF-16 unit, as in Find: a character outside the BMP is two.
 */
function matcher(query: string, o: FindCellsOptions): ((text: string) => boolean) | null {
  const fold = o.caseSensitive ? (t: string) => t : folded
  const q = fold(query)
  if (!o.wildcards) return q ? (t) => (o.wholeCell ? fold(t) === q : fold(t).includes(q)) : null
  const source = wildSource(q)
  if (!source) return null
  const re = new RegExp(o.wholeCell ? `^(?:${source})$` : source)
  return (t) => t !== '' && re.test(fold(t))
}

/** A sheet's hidden rows or columns: whether there are any; whether a line is hidden, and how many lines are hidden before it, the band of lines it is in. */
interface HiddenLines {
  any: boolean
  at(i: number): [hidden: boolean, before: number]
}

/** A sheet's hidden rows or columns. */
function hiddenLines<T>(map: Record<string, T> | undefined, axis: 'r' | 'c', hidden: (e: T) => boolean): HiddenLines {
  const spans: [number, number][] = []
  for (const [lo, hi, e] of lines(map, axis)) {
    if (!hidden(e)) continue
    const last = spans.at(-1)
    if (last && last[1] + 1 === lo) last[1] = hi
    else spans.push([lo, hi])
  }
  // the lines hidden before each span, and in all
  const before = [0]
  for (const [lo, hi] of spans) before.push(before.at(-1)! + hi - lo + 1)
  return {
    any: spans.length > 0,
    at(i) {
      const k = lowerBound(spans, i, (x) => x[1])
      return k < spans.length && spans[k][0] <= i ? [true, 0] : [false, before[k]]
    },
  }
}

/** A finder: its options as a cursor's key, the cells in its order and the test of each. */
export interface Finder {
  key: unknown[]
  /** The cells of `refs` from `cells` (each once, in the first range that has it), range by range in Find's order. */
  cells(refs: Ref[], cells: Iterable<At | undefined>): Generator<At | void>
  matches(s: number, r: number, c: number): boolean
}

export function finderOf(state: EngineState, query: string, o: FindCellsOptions, now: string): Finder {
  const where = o.in ?? 'text'
  if (!WHERE.includes(where)) throw new InputError('opts.in', 'expected text, edit, formula or value')
  for (const k of FIND_OPTIONS.slice(1)) if (o[k] !== undefined) checkBoolean(o[k], `opts.${k}`)
  const test = matcher(query, o)
  const looks = where === 'text' ? cellLooks(state) : null
  const perSheet = new Map<number, { lines: SheetLines; edit: ReturnType<typeof editTexts>; rows: HiddenLines; cols: HiddenLines }>()
  const sheet = (s: number) => {
    let x = perSheet.get(s)
    if (!x) {
      const input = state.input.sheets[s]
      perSheet.set(s, (x = {
        lines: sheetLines(input, state),
        edit: editTexts(state.input, s, state.tables, now),
        rows: hiddenLines(input.rows, 'r', isHiddenRow),
        cols: hiddenLines(input.columns, 'c', (e) => e.hidden === true || e.width === 0),
      }))
    }
    return x
  }
  // what the sheet's conditional formats give the cells of the range being read: the text by a format of theirs
  let shown: ReturnType<typeof appearanceOf> | null = null

  // a cell's place by rows (row-major) or by columns, and back
  const place = (r: number, c: number) => (o.byColumns ? c * MAX_ROWS + r : r * MAX_COLS + c)
  const cell = (m: number): [number, number] => (o.byColumns ? [m % MAX_ROWS, Math.floor(m / MAX_ROWS)] : [Math.floor(m / MAX_COLS), m % MAX_COLS])

  /** A range's cells in Find's order: by their places, after the blocks hidden lines leave (by rows of blocks, then by columns) when hidden cells are skipped. */
  const ordered = (s: number, places: number[]): Float64Array => {
    const xs = Float64Array.from(places).sort()
    const { rows, cols } = sheet(s)
    if (!o.skipHidden || (!rows.any && !cols.any)) return xs
    const block = xs.map((m) => {
      const [r, c] = cell(m)
      return rows.at(r)[1] * MAX_COLS + cols.at(c)[1]
    })
    const ix = Array.from(xs.keys()).sort((a, b) => block[a] - block[b] || xs[a] - xs[b])
    return Float64Array.from(ix, (i) => xs[i])
  }

  function* cells(refs: Ref[], all: Iterable<At | undefined>): Generator<At | void> {
    if (!test) return
    const groups = refs.map((): number[] => [])
    for (const at of all) {
      if (!at) {
        yield
        continue
      }
      const [s, r, c] = at
      if (o.skipHidden && (sheet(s).rows.at(r)[0] || sheet(s).cols.at(c)[0])) continue
      groups[refs.findIndex((x) => x.s === s && hasCell(x, r, c))].push(place(r, c))
    }
    for (const [n, group] of groups.entries()) {
      const s = refs[n].s
      const places = ordered(s, group)
      shown = null
      if (where === 'text' && state.input.sheets[s].conditionalFormats?.length) {
        const only = new Set(Array.from(places, (m) => cell(m)).map(([r, c]) => r * MAX_COLS + c))
        shown = appearanceOf(yield* windowLooks(state, s, refs[n], only), looks, s)
      }
      for (const m of places) {
        if (paced(1)) yield
        yield [s, ...cell(m)]
      }
    }
  }

  /** The text Find reads in a cell. */
  const textOf = (s: number, r: number, c: number): string => {
    const own = state.input.sheets[s].cells[formatAddress(r, c)]
    const id = where === 'text' || where === 'value' || own?.formula === undefined ? undefined : state.at.get(key(s, r, c))
    const formula = id === undefined ? own?.formula : state.formulas[id]!.text
    if (where === 'formula') return formula ?? ''
    const value = outValue(toCellValue(valueAt(state, s, r, c)))
    if (looks) {
      const view = { value, ...looks(s, r, c, value, own ?? {}) }
      shown?.(view, r, c, own)
      return view.text ?? ''
    }
    if (where === 'value') return valueText(value)
    // Look in Formulas reads a text as it is, without the apostrophe the formula bar shows before one that reads as a number
    if (formula === undefined && own && typeof value === 'string') return value
    const x = own ? undefined : ownerOf(state, s, r, c)
    const { lines, edit } = sheet(s)
    return edit.editText({ formula, value, ...(x ? { spillOf: formatAddress(x.r1, x.c1) } : {}) }, own, lines.format(r, c, own), r, c)
  }

  return {
    key: [query, where, o.caseSensitive === true, o.wildcards === true, o.wholeCell === true, o.byColumns === true, o.skipHidden === true],
    cells,
    matches: (s, r, c) => test!(textOf(s, r, c)),
  }
}

const valueText = (v: CellValue): string => (typeof v === 'object' && v !== null ? JSON.stringify(v) : String(v ?? ''))
