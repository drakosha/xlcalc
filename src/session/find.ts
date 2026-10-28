/**
 * Excel's Find in a session (`Session.findCells`), measured in Excel for Windows (`bench/coverage/find/`): the query read
 * as it is or with Find's wildcards, matched against a part or the whole of what a cell shows (Look in Values), the text
 * it opens with for editing (Look in Formulas), its stored formula or its value, in any case or in its own; the cells of
 * each range by rows or by columns, those of hidden rows and columns skipped as Find skips them. The texts are read once
 * and kept with the session's state (`search.ts`), so another search of the same sheet only matches them.
 */
import type { CellValue, ColumnInput, FindCellsOptions, RowInput } from '../contract.js'
import type { EngineState } from '../engine.js'
import { InputError } from '../input-error.js'
import { ERROR_TEXT } from '../syntax/errors.js'
import { checkBoolean } from '../input/basic.js'
import { hiddenSpansOf } from '../layout/hidden.js'
import { colOf, formatAddress, MAX_COLS, MAX_ROWS, placeOf, rowOf } from '../model.js'
import { cellValue } from '../result.js'
import { cellLooks } from '../shown.js'
import { appearanceOf, windowLooks } from '../cond/looks.js'
import { hasCell, meets, type Rect } from '../structure/area.js'
import { paced, type Steps } from '../exec/budget.js'
import { folded, matcher } from '../find-text.js'
import { lowerBound } from '../sorted.js'
import { FIND_OPTIONS } from '../strict-args.js'
import { inputIn } from '../looks/sheet.js'
import { sessionEditTexts } from './viewport.js'
import { byColumns, kindOf, textsOf, type SheetTexts } from './search.js'

const WHERE = ['text', 'edit', 'formula', 'value'] as const

export type Ref = Rect & { s: number }

/** A cell found: its sheet, row and column, its position in Find's order (its range, its block's bands of rows and columns, its place), which a cursor keeps, and the text Find read in it. */
export type Found = [s: number, r: number, c: number, position: number[], text: string]

/** A sheet's hidden rows or columns: whether there are any; whether a line is hidden, and how many lines are hidden before it, the band of lines it is in (a hidden line's, the band after it). */
interface HiddenLines {
  any: boolean
  at(i: number): [hidden: boolean, before: number]
}

/** A sheet's hidden rows or columns. */
function hiddenLines(map: Record<string, RowInput | ColumnInput> | undefined, axis: 'r' | 'c'): HiddenLines {
  const spans = hiddenSpansOf(map, axis)
  // the lines hidden before each span, and in all
  const before = [0]
  for (const [lo, hi] of spans) before.push(before.at(-1)! + hi - lo + 1)
  return {
    any: spans.length > 0,
    at(i) {
      const k = lowerBound(spans, i, (x) => x[1])
      return k < spans.length && spans[k][0] <= i ? [true, before[k + 1]] : [false, before[k]]
    },
  }
}

/** A finder: its options as a cursor's key, and the cells it finds in Find's order. */
export interface Finder {
  key: unknown[]
  /**
   * The cells of `refs` that match (each once, in the first range that has it), range by range in Find's order, after
   * the position `from` of a cell found before and up to the position `to`; with `rows` (zero-based, ascending), only
   * the cells of those rows.
   */
  found(refs: Ref[], rows: number[] | null, from: number[] | null, to?: number[] | null): Generator<Found | void>
  around(refs: Ref[], rows: number[] | null, start: number[], cursor: number[] | null): Generator<Found | void>
  positionOf(refs: Ref[], s: number, r: number, c: number): number[] | null
}

export function finderOf(state: EngineState, query: string, o: FindCellsOptions, now: string): Finder {
  const where = o.in ?? 'text'
  if (!WHERE.includes(where)) throw new InputError('opts.in', 'expected text, edit, formula or value')
  for (const k of FIND_OPTIONS.slice(1)) if (o[k] !== undefined) checkBoolean(o[k], `opts.${k}`)
  const test = matcher(query, o)
  const fold = !o.caseSensitive
  const looks = where === 'text' ? cellLooks(state) : null
  const editText = sessionEditTexts(state, now)
  const hidden = new Map<number, { rows: HiddenLines; cols: HiddenLines }>()
  const linesOf = (s: number) => {
    let x = hidden.get(s)
    if (!x) {
      const input = state.input.sheets[s]
      hidden.set(s, (x = { rows: hiddenLines(input.rows, 'r'), cols: hiddenLines(input.columns, 'c') }))
    }
    return x
  }

  // Find skips hidden cells in Look in Values, and on a filtered sheet in every kind of search, where it takes the cells by blocks
  const filtered = (s: number) => o.skipHidden === true && state.input.sheets[s].filtered === true
  const skips = (s: number) => filtered(s) || (o.skipHidden === true && where === 'text')

  /** The text Find reads in a cell; `shown`, what the sheet's conditional formats give the cells being read. */
  const textOf = (s: number, r: number, c: number, shown: ReturnType<typeof appearanceOf> | null): string => {
    // a cell of the sheet's looks reads as an entry of its look alone
    const own = inputIn(state.input.sheets[s], formatAddress(r, c), r, c)
    // Look in Formulas reads a text as it is, without the apostrophe the formula bar shows before one that reads as a number
    if (where === 'edit') return own?.formula === undefined && typeof own?.value === 'string' ? own.value : editText(s, r, c, own)
    if (where === 'formula') {
      const id = own?.formula === undefined ? undefined : state.at.get(placeOf(s, r, c))
      return (id === undefined ? own?.formula : state.formulas[id]!.text) ?? ''
    }
    const value = cellValue(state, s, r, c)
    if (!looks) return valueText(value)
    const view = { value, ...looks(s, r, c, value, own ?? {}) }
    shown?.(view, r, c, own)
    return view.text ?? ''
  }

  /** The texts of cells `ix` of a sheet that none has been read of yet, read: in their case (`raw`), and folded into `texts` without Match case. */
  function* fill(s: number, t: SheetTexts, ix: number[], raw: (string | undefined)[], texts: (string | undefined)[]): Steps<void> {
    let shown: ReturnType<typeof appearanceOf> | null = null
    const need = ix.filter((i) => raw[i] === undefined)
    if (looks && need.length && state.input.sheets[s].conditionalFormats?.length) {
      const x = { r1: MAX_ROWS, c1: MAX_COLS, r2: 0, c2: 0 }
      for (const i of need) {
        const [r, c] = rc(t.at[i])
        x.r1 = Math.min(x.r1, r)
        x.r2 = Math.max(x.r2, r)
        x.c1 = Math.min(x.c1, c)
        x.c2 = Math.max(x.c2, c)
      }
      shown = appearanceOf(yield* windowLooks(state, s, x, new Set(need.map((i) => t.at[i]))), looks, s)
    }
    for (const i of ix) {
      if (paced(1)) yield
      raw[i] ??= textOf(s, ...rc(t.at[i]), shown)
      texts[i] = fold ? folded(raw[i]) : raw[i]
    }
  }

  /**
   * The cells of range `n` that may match, in their order after the place `from`, a chunk at a time: their indices in
   * the sheet's texts and their places in that order. The steps yield between cells.
   */
  function* chunksOf(refs: Ref[], n: number, t: SheetTexts, rows: number[] | null, from: number | null): Generator<[number[], number[]] | void> {
    const x = refs[n]
    const { rows: hr, cols: hc } = linesOf(x.s)
    const skip = skips(x.s) && (hr.any || hc.any)
    const earlier = refs.slice(0, n).filter((y) => y.s === x.s && meets(y, x))
    let ix: number[] = []
    let places: number[] = []
    const take = (i: number, place: number, r: number, c: number) => {
      if (skip && (hr.at(r)[0] || hc.at(c)[0])) return false
      if (earlier.length && earlier.some((y) => hasCell(y, r, c))) return false
      ix.push(i)
      places.push(place)
      return ix.length === CHUNK
    }
    const start = from === null ? 0 : from + 1
    // by columns: the range's columns at once; by rows: the rows asked of the range one by one, or its rows at once
    const cols = o.byColumns ? byColumns(t) : null
    const only = cols && rows ? new Set(rows) : null
    const spans: [number, number][] = cols ? [[x.c1 * MAX_ROWS + x.r1, x.c2 * MAX_ROWS + x.r2]] : (rows ? rows.slice(lowerBound(rows, x.r1), lowerBound(rows, x.r2 + 1)) : [x.r1]).map((r) => [r * MAX_COLS + x.c1, (rows ? r : x.r2) * MAX_COLS + x.c2])
    const at = cols ? cols.at : t.at
    for (const [lo, hi] of spans)
      for (let j = lowerBound(at, Math.max(lo, start)); j < at.length && at[j] <= hi; j++) {
        if (paced(1)) yield
        const k = at[j]
        // the row and column of the place, without a pair made for each cell
        const q = Math.floor(k / (cols ? MAX_ROWS : MAX_COLS))
        const r = cols ? k - q * MAX_ROWS : q
        const c = cols ? q : k - q * MAX_COLS
        if (r < x.r1 || r > x.r2 || c < x.c1 || c > x.c2 || (only && !only.has(r)) || !take(cols ? cols.ix[j] : j, k, r, c)) continue
        yield [ix, places]
        ix = []
        places = []
      }
    if (ix.length) yield [ix, places]
  }

  /**
   * A cell's position in Find's order over `refs`, as `found` gives one; null outside them. On a filtered sheet a cell of
   * a hidden line stands in its range's first block, as Excel's Find from it goes on there (the set `find/rules-held-4`).
   */
  const positionOf = (refs: Ref[], s: number, r: number, c: number): number[] | null => {
    const n = refs.findIndex((x) => x.s === s && hasCell(x, r, c))
    if (n < 0) return null
    const { rows: hr, cols: hc } = linesOf(s)
    const place = o.byColumns ? c * MAX_ROWS + r : r * MAX_COLS + c
    if (!filtered(s) || (!hr.any && !hc.any)) return [n, 0, 0, place]
    const hidden = hr.at(r)[0] || hc.at(c)[0]
    return [n, hr.at(hidden ? refs[n].r1 : r)[1], hc.at(hidden ? refs[n].c1 : c)[1], place]
  }

  function* found(refs: Ref[], rows: number[] | null, from: number[] | null, to: number[] | null = null): Generator<Found | void> {
    if (!test) return
    for (const [n, x] of refs.entries()) {
      if ((from && n < from[0]) || (to && n > to[0])) continue
      const t = yield* textsOf(state, x.s)
      const { rows: hr, cols: hc } = linesOf(x.s)
      // on a filtered sheet with hidden lines the cells come by blocks, all of the range's put in order: the block's bands first, a cursor's too
      const blocks = filtered(x.s) && (hr.any || hc.any)
      const [after, until] = [from && n === from[0] ? from : null, to && n === to[0] ? to : null]
      const raw = kindOf(t, where)
      const texts = fold ? kindOf(t, `${where} folded`) : raw
      const hits: number[][] = []
      for (const chunk of chunksOf(refs, n, t, rows, blocks || !after ? null : after[3])) {
        if (!chunk) {
          yield
          continue
        }
        const [ix, places] = chunk
        const unread = ix.filter((i) => texts[i] === undefined)
        if (unread.length) yield* fill(x.s, t, unread, raw, texts)
        for (let q = 0; q < ix.length; q++) {
          if (paced(1)) yield
          const i = ix[q]
          if (!test(texts[i]!)) continue
          const [r, c] = rc(t.at[i])
          const position = [n, blocks ? hr.at(r)[1] : 0, blocks ? hc.at(c)[1] : 0, places[q]]
          if (blocks) hits.push([...position, i])
          else if (until && later(position, until)) return
          else yield [x.s, r, c, position, raw[i]!]
        }
      }
      hits.sort((a, b) => a[1] - b[1] || a[2] - b[2] || a[3] - b[3])
      for (const [, rb, cb, place, i] of hits) {
        const position = [n, rb, cb, place]
        if ((!after || later(position, after)) && (!until || !later(position, until))) yield [x.s, ...rc(t.at[i]), position, raw[i]!]
      }
    }
  }

  /**
   * The cells `found` gives, from after the position `start` round to it (Find from a cell, then FindNext), each with its
   * lap (0 after `start`, 1 from the first) before its position; after `cursor`, a position so given.
   */
  function* around(refs: Ref[], rows: number[] | null, start: number[], cursor: number[] | null): Generator<Found | void> {
    const lap = cursor?.[0] ?? 0
    for (const [k, from, to] of [[0, lap ? null : (cursor?.slice(1) ?? start), null], [1, lap ? cursor!.slice(1) : null, start]] as const) {
      if (k < lap) continue
      for (const f of found(refs, rows, from, to)) yield f && [f[0], f[1], f[2], [k, ...f[3]], f[4]]
    }
  }

  return {
    key: [query, where, o.caseSensitive === true, o.wildcards === true, o.wholeCell === true, o.byColumns === true, o.skipHidden === true],
    found,
    around,
    positionOf,
  }
}

/** A place's row and column. */
const rc = (m: number): [number, number] => [rowOf(m), colOf(m)]
/** Whether position `a` comes after `b` in Find's order. */
const later = (a: number[], b: number[]) => {
  for (let k = 0; k < a.length; k++) if (a[k] !== b[k]) return a[k] > b[k]
  return false
}

/** The cells a search reads at a time, so that a page that is full stops it. */
const CHUNK = 4096

/** A value as the `value` kind reads it, the same in every locale: a number as JavaScript writes it, TRUE or FALSE, an error by its name. */
const valueText = (v: CellValue): string => (typeof v === 'object' && v !== null ? ERROR_TEXT[v.error] : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v ?? ''))
