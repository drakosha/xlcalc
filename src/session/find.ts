/**
 * Excel's Find in a session (`Session.findCells`), measured in Excel for Windows (`bench/coverage/find/`): the query read
 * as it is or with Find's wildcards, matched against a part or the whole of what a cell shows (Look in Values), the text
 * it opens with for editing (Look in Formulas), its stored formula or its value, in any case or in its own; the cells of
 * each range by rows or by columns, those of hidden rows and columns skipped as Find skips them. The texts are read once
 * and kept with the session's state (`search.ts`), so another search of the same sheet only matches them.
 */
import type { CellValue, FindCellsOptions } from '../contract.js'
import { literal, wildSource } from '../criteria.js'
import { key, toCellValue, valueAt, type EngineState } from '../engine.js'
import { checkBoolean, InputError } from '../input.js'
import { isHiddenRow } from '../layout/hidden.js'
import { lines } from '../layout/spans.js'
import { formatAddress, MAX_COLS, MAX_ROWS } from '../model.js'
import { outValue } from '../result.js'
import { cellLooks } from '../shown.js'
import { appearanceOf, windowLooks } from '../cond/looks.js'
import { hasCell, meets, type Rect } from '../structure/area.js'
import { paced, type Steps } from '../exec/budget.js'
import { lowerBound } from '../sorted.js'
import { FIND_OPTIONS } from '../strict-args.js'
import { inputIn } from '../looks/sheet.js'
import { sessionEditTexts } from './viewport.js'
import { byColumns, kindOf, textsOf, type SheetTexts } from './search.js'

const WHERE = ['text', 'edit', 'formula', 'value'] as const

export type Ref = Rect & { s: number }

/** A cell found: its sheet, row and column, and its position in Find's order (its range, its block, its place), which a cursor keeps. */
export type Found = [s: number, r: number, c: number, position: number[]]

const same = (v: number) => v

/**
 * Each character's lower case as one UTF-16 unit or two as it was, so `?` keeps its place: `İ` stays (its lower case is
 * two characters), and the final `ς` is `σ`, as `Σ` is (Find takes `σ`, `ς` and `Σ` as one letter without Match case).
 */
const folded = (t: string) => {
  const lower = t.toLowerCase()
  const f = lower.length === t.length ? lower : t.replace(/[^İ]+/g, (m) => m.toLowerCase())
  return f.includes('ς') ? f.replaceAll('ς', 'σ') : f
}

/**
 * The test of a cell's text (folded without Match case), null when nothing can match (an empty query, or `~` alone).
 * Wildcards as the lookup functions read them (`wildSource`) but `?` is one UTF-16 unit, as in Find: a character
 * outside the BMP is two.
 */
function matcher(query: string, o: FindCellsOptions): ((text: string) => boolean) | null {
  const q = o.caseSensitive ? query : folded(query)
  if (!o.wildcards) return q ? (t) => (o.wholeCell ? t === q : t.includes(q)) : null
  const source = wildSource(q)
  if (!source) return null
  const re = new RegExp(o.wholeCell ? `^(?:${source})$` : source)
  return (t) => t !== '' && re.test(t)
}

/**
 * Excel's Replace in a text: each match of the query, as `matcher` reads it, put as `by` is, the whole text with Match
 * entire cell contents; null where nothing matches or nothing can (an empty query). An empty match, of `*` at the end,
 * puts nothing.
 */
export function replacer(query: string, by: string, o: Pick<FindCellsOptions, 'caseSensitive' | 'wildcards' | 'wholeCell'>): ((text: string) => string | null) | null {
  const q = o.caseSensitive ? query : folded(query)
  const source = o.wildcards ? wildSource(q) : [...q].map(literal).join('')
  if (!source) return null
  const re = new RegExp(o.wholeCell ? `^(?:${source})$` : source, 'g')
  return (text) => {
    let out = ''
    let last = 0
    for (const m of (o.caseSensitive ? text : folded(text)).matchAll(re)) {
      if (m[0] === '') continue
      out += text.slice(last, m.index) + by
      last = m.index + m[0].length
    }
    return last ? out + text.slice(last) : null
  }
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

/** A finder: its options as a cursor's key, and the cells it finds in Find's order. */
export interface Finder {
  key: unknown[]
  /**
   * The cells of `refs` that match (each once, in the first range that has it), range by range in Find's order, after
   * the position `after` of a cell found before; with `rows` (zero-based, ascending), only the cells of those rows.
   */
  found(refs: Ref[], rows: number[] | null, after: number[] | null): Generator<Found | void>
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
      hidden.set(s, (x = { rows: hiddenLines(input.rows, 'r', isHiddenRow), cols: hiddenLines(input.columns, 'c', (e) => e.hidden === true || e.width === 0) }))
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
      const id = own?.formula === undefined ? undefined : state.at.get(key(s, r, c))
      return (id === undefined ? own?.formula : state.formulas[id]!.text) ?? ''
    }
    const value = outValue(toCellValue(valueAt(state, s, r, c)))
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
    const spans: [number, number][] = cols ? [[x.c1 * MAX_ROWS + x.r1, x.c2 * MAX_ROWS + x.r2]] : (rows ? rows.slice(lowerBound(rows, x.r1, same), lowerBound(rows, x.r2 + 1, same)) : [x.r1]).map((r) => [r * MAX_COLS + x.c1, (rows ? r : x.r2) * MAX_COLS + x.c2])
    const at = cols ? cols.at : t.at
    for (const [lo, hi] of spans)
      for (let j = lowerBound(at, Math.max(lo, start), same); j < at.length && at[j] <= hi; j++) {
        if (paced(1)) yield
        const k = at[j]
        const [r, c] = cols ? cr(k) : rc(k)
        if (r < x.r1 || r > x.r2 || c < x.c1 || c > x.c2 || (only && !only.has(r)) || !take(cols ? cols.ix[j] : j, k, r, c)) continue
        yield [ix, places]
        ix = []
        places = []
      }
    if (ix.length) yield [ix, places]
  }

  function* found(refs: Ref[], rows: number[] | null, after: number[] | null): Generator<Found | void> {
    if (!test) return
    for (const [n, x] of refs.entries()) {
      if (after && n < after[0]) continue
      const t = yield* textsOf(state, x.s)
      const { rows: hr, cols: hc } = linesOf(x.s)
      // on a filtered sheet with hidden lines the cells come by blocks, all of the range's put in order: the block's place first, a cursor's too
      const blocks = filtered(x.s) && (hr.any || hc.any)
      const from = after && n === after[0] ? after : null
      const raw = kindOf(t, where)
      const texts = fold ? kindOf(t, `${where} folded`) : raw
      const hits: [block: number, place: number, i: number][] = []
      for (const chunk of chunksOf(refs, n, t, rows, blocks || !from ? null : from[2])) {
        if (!chunk) {
          yield
          continue
        }
        const [ix, places] = chunk
        const unread = ix.filter((i) => texts[i] === undefined)
        if (unread.length) yield* fill(x.s, t, unread, raw, texts)
        for (const [q, i] of ix.entries()) {
          if (paced(1)) yield
          if (!test(texts[i]!)) continue
          if (!blocks) {
            yield [x.s, ...rc(t.at[i]), [n, 0, places[q]]]
            continue
          }
          const [r, c] = rc(t.at[i])
          const block = hr.at(r)[1] * MAX_COLS + hc.at(c)[1]
          if (!from || block > from[1] || (block === from[1] && places[q] > from[2])) hits.push([block, places[q], i])
        }
      }
      hits.sort((a, b) => a[0] - b[0] || a[1] - b[1])
      for (const [block, place, i] of hits) yield [x.s, ...rc(t.at[i]), [n, block, place]]
    }
  }

  return {
    key: [query, where, o.caseSensitive === true, o.wildcards === true, o.wholeCell === true, o.byColumns === true, o.skipHidden === true],
    found,
  }
}

/** A place's row and column; by columns (`c * MAX_ROWS + r`), too. */
const rc = (m: number): [number, number] => {
  const r = Math.floor(m / MAX_COLS)
  return [r, m - r * MAX_COLS]
}
const cr = (k: number): [number, number] => {
  const c = Math.floor(k / MAX_ROWS)
  return [k - c * MAX_ROWS, c]
}

/** The cells a search reads at a time, so that a page that is full stops it. */
const CHUNK = 4096

const valueText = (v: CellValue): string => (typeof v === 'object' && v !== null ? JSON.stringify(v) : String(v ?? ''))
