/**
 * Excel's fill handle (`StructureOp` `fillRange`, `Range.AutoFill`). The cells added take the source's cells in turn as
 * copies of it: whole sources next to it, then the part of it nearest to it at the far end (`copyRange` each, so formulas,
 * looks, rules and validations go as a copy takes them). Then, unless the fill copies alone, each line along the fill (a
 * column of a fill down or up, a row of one right or left) is cut into runs of its source's cells that go on together, and
 * the copies of a run's cells take the run's next values: numbers by their trend, dates by days or by months, times by
 * hours, text that ends in a number by the number, the names of days and months through their list. A run's cells come
 * again in every copy of the source, so the run goes on from copy to copy (1, 2, x → 3, 4, x, 5, 6, x), and backwards in a
 * fill up or left.
 */
import type { BookInput, CellInput, FillType } from '../contract.js'
import { addMonths, daysInMonth, epochOf, serialOf, ymd } from '../dates.js'
import { dateToSerial, serialToDate } from '../date-serial.js'
import { InputError } from '../input.js'
import { lineAt } from '../layout/spans.js'
import { localeOf } from '../locale/registry.js'
import { formatAddress } from '../model.js'
import { EN_BUILTIN } from '../numfmt/builtin.js'
import { minutes } from '../numfmt/date.js'
import { EN_CODES, parseFormat } from '../numfmt/parse.js'
import { EN_NAMES } from '../numfmt/style.js'
import { hi, lo, withBounds, type Axis, type Rect } from './area.js'
import type { Applied, Geometry } from './geometry.js'
import { cellsIn, gridOf } from './grid.js'
import { sheetIndex } from './rewrite.js'

/** A copy of part of the sheet over another part, as `editStructure` applies one (`copyRange`). */
export type Copier = (book: BookInput, from: Rect, to: Rect) => Applied

/** A cell of a line's source as a fill reads it: what kind of series it can go on in, and its value there. */
type Item =
  | { k: 'number' | 'date' | 'time'; y: number }
  /** Text that goes on by a number in it: the cells of one `key` go on together, `make` writes the text of another number. */
  | { k: 'text'; key: string; y: number; make: (n: number) => string }
  | { k: 'other' }

/** A run of a line's source: consecutive cells that go on together, from position `at` of the source. */
interface Run {
  at: number
  items: Item[]
}

/**
 * The fill of `src` over `target` on `sheet` (which holds `src` and reaches past it along one axis), its copies made by
 * `copy`; `to` is the target as given, for the message of one of the wrong shape.
 */
export function fillRange(book: BookInput, sheet: string, src: Rect, target: Rect, to: string, type: FillType, copy: Copier): Applied {
  const way = wayOf(src, target)
  if (!way) throw new InputError('op.to', `expected a range that holds op.range and reaches past it down, up, right or left, got ${JSON.stringify(to)}`)
  const { axis, forward } = way
  const [a, b] = [lo(src, axis), hi(src, axis)]
  const h = b - a + 1
  const n = hi(target, axis) - lo(target, axis) + 1 - h
  const [whole, part] = [n - (n % h), n % h]
  const pieces: [Rect, Rect][] = []
  if (forward) {
    if (whole) pieces.push([src, withBounds(src, axis, b + 1, b + whole)])
    if (part) pieces.push([withBounds(src, axis, a, a + part - 1), withBounds(src, axis, b + whole + 1, b + whole + part)])
  } else {
    if (whole) pieces.push([src, withBounds(src, axis, a - whole, a - 1)])
    if (part) pieces.push([withBounds(src, axis, b - part + 1, b), withBounds(src, axis, a - whole - part, a - whole - 1)])
  }
  let out = book
  const geometry: Geometry = []
  for (const [from, at] of pieces) {
    const r = copy(out, from, at)
    if (!r.ok) return r
    out = r.book
    geometry.push(...r.geometry)
  }
  if (type === 'copy') return { ok: true, book: out, geometry }
  const s = sheetIndex(book, sheet)
  const written = seriesCells(book, s, src, axis, forward ? [b + 1, b + n] : [a - n, a - 1], type)
  if (!written.length) return { ok: true, book: out, geometry }
  const cells = { ...out.sheets[s].cells }
  for (const [addr, value] of written) cells[addr] = { ...cells[addr], value }
  return { ok: true, book: { ...out, sheets: out.sheets.map((x, i) => (i === s ? { ...x, cells } : x)) }, geometry }
}

/** The axis and the way a fill goes: `target` holds `src`, has its width across, and reaches past one of its ends. */
function wayOf(src: Rect, target: Rect): { axis: Axis; forward: boolean } | null {
  for (const axis of ['r', 'c'] as const) {
    const across: Axis = axis === 'r' ? 'c' : 'r'
    if (lo(target, across) !== lo(src, across) || hi(target, across) !== hi(src, across)) continue
    if (lo(target, axis) === lo(src, axis) && hi(target, axis) > hi(src, axis)) return { axis, forward: true }
    if (hi(target, axis) === hi(src, axis) && lo(target, axis) < lo(src, axis)) return { axis, forward: false }
  }
  return null
}

/** The cells of positions `added` (from, to along the axis) that a series writes, with their values. */
function seriesCells(book: BookInput, s: number, src: Rect, axis: Axis, added: [number, number], type: FillType): [string, CellInput['value']][] {
  const sheet = book.sheets[s]
  const g = gridOf(sheet)
  const a = lo(src, axis)
  const h = hi(src, axis) - a + 1
  // the source's cells by line across the fill, at their position along it
  const lines = new Map<number, Map<number, CellInput>>()
  for (const i of cellsIn(g, src)) {
    const [along, across] = axis === 'r' ? [g.row[i], g.col[i]] : [g.col[i], g.row[i]]
    let line = lines.get(across)
    if (!line) lines.set(across, (line = new Map()))
    line.set(along - a, g.cell[i])
  }
  const lists = listsOf(book)
  const epoch = epochOf(book.dateSystem)
  const out: [string, CellInput['value']][] = []
  for (const [across, line] of lines) {
    const items = Array.from({ length: h }, (_, p) => {
      const at = a + p
      const [r, c] = axis === 'r' ? [at, across] : [across, at]
      return itemOf(line.get(p), formatAt(sheet, line.get(p), r, c), book, lists)
    })
    // a number alone in its line is copied as a single cell is, unless the fill is a series
    const alone = line.size === 1 || [...line.values()].filter((x) => x.formula !== undefined || (x.value !== undefined && x.value !== null)).length === 1
    for (const run of runsOf(items)) {
      const next = goesOn(run, type === 'series' || !alone, epoch)
      if (!next) continue
      const m = run.items.length
      for (let pos = added[0]; pos <= added[1]; pos++) {
        const p = pos - a
        const q = ((p % h) + h) % h
        if (q < run.at || q >= run.at + m) continue
        const e = Math.floor(p / h) * m + (q - run.at)
        const [r, c] = axis === 'r' ? [pos, across] : [across, pos]
        const source = run.items[q - run.at]
        const value = next(e, source)
        const own = line.get(q)!.value
        out.push([formatAddress(r, c), typeof own === 'object' && own !== null && typeof value === 'number' ? { date: serialToDate(value, { dateSystem: book.dateSystem }) } : value])
      }
    }
  }
  return out
}

/** A cell's number format: its own, else its row's, else its column's. */
function formatAt(sheet: BookInput['sheets'][number], cell: CellInput | undefined, r: number, c: number) {
  return cell?.numberFormat ?? lineAt(sheet.rows, 'r', r)?.numberFormat ?? lineAt(sheet.columns, 'c', c)?.numberFormat
}

/** The lists a text goes on through: en-US's short and long day and month names, and the book's locale's. */
function listsOf(book: BookInput): (readonly string[])[] {
  const en = [EN_NAMES.daysShort, EN_NAMES.days, EN_NAMES.monthsShort, EN_NAMES.months]
  const r = localeOf(book.locale).regional
  return r && r.lcid !== '409' ? [...en, r.daysShort, r.days, r.monthsShort, r.months] : en
}

/** Text that ends in a number: the text before it and its digits. */
const SUFFIX = /^(.*?)(\d+)$/s
/** Text that starts with a number: its digits and the text after them. */
const LEADING = /^(\d+)(\D.*)$/s
/** A quarter: `Q1`, `Qtr 2`, `Quarter 4`. */
const QUARTER = /^(q|qtr|quarter)( ?)([1-4])$/i
/** An ordinal of en-US: `1st`, `22nd`. */
const ORDINAL = /^(\d+)(st|nd|rd|th)$/i

/** `n` within 0 … `size` - 1, as a list goes round. */
const round = (n: number, size: number) => ((n % size) + size) % size

/** A number's digits, as wide as `width` with zeros before them; a number below zero by its size. */
const digits = (n: number, width: number) => String(Math.abs(n)).padStart(width, '0')

/** The width a number's digits keep: theirs when they start with a zero, else none. */
const widthOf = (d: string) => (d.length > 1 && d[0] === '0' ? d.length : 0)

/** en-US's ending of an ordinal: 1st, 2nd, 3rd, 4th, 11th, 21st. */
function ordinal(n: number): string {
  const [t, u] = [Math.abs(n) % 100, Math.abs(n) % 10]
  return t >= 11 && t <= 13 ? 'th' : u === 1 ? 'st' : u === 2 ? 'nd' : u === 3 ? 'rd' : 'th'
}

/** How a fill reads a cell of its source. */
function itemOf(cell: CellInput | undefined, format: CellInput['numberFormat'], book: BookInput, lists: (readonly string[])[]): Item {
  if (!cell || cell.formula !== undefined) return { k: 'other' }
  const v = cell.value
  if (typeof v === 'object' && v !== null) {
    const serial = dateSerial(v.date, book)
    return serial === null ? { k: 'other' } : { k: 'date', y: serial }
  }
  if (typeof v === 'number') return { k: shownAs(format), y: v }
  if (typeof v !== 'string') return { k: 'other' }
  const lower = v.toLowerCase()
  for (const [i, list] of lists.entries()) {
    const y = list.findIndex((x) => x.toLowerCase() === lower)
    const caseOf = casing(v)
    if (y >= 0) return { k: 'text', key: `list ${i}`, y, make: (n) => caseOf(list[round(n, list.length)]) }
  }
  let m = QUARTER.exec(v)
  if (m) {
    const head = m[1] + m[2]
    return { k: 'text', key: `quarter ${head.toLowerCase()}`, y: Number(m[3]), make: (n) => head + (round(n - 1, 4) + 1) }
  }
  if ((m = ORDINAL.exec(v))) {
    const caseOf = casing(m[2])
    return { k: 'text', key: 'ordinal', y: Number(m[1]), make: (n) => Math.abs(n) + caseOf(ordinal(n)) }
  }
  if ((m = SUFFIX.exec(v))) {
    const [head, width] = [m[1], widthOf(m[2])]
    return { k: 'text', key: `suffix ${head}`, y: Number(m[2]), make: (n) => head + digits(n, width) }
  }
  if ((m = LEADING.exec(v))) {
    const [tail, width] = [m[2], widthOf(m[1])]
    return { k: 'text', key: `leading ${tail}`, y: Number(m[1]), make: (n) => digits(n, width) + tail }
  }
  return { k: 'other' }
}

/** What a number's format shows it as, by its first section: a date (a day, a month or a year in it), a time of day alone, or a number. */
function shownAs(format: CellInput['numberFormat']): 'date' | 'time' | 'number' {
  if (format === undefined) return 'number'
  const toks = parseFormat(typeof format === 'number' ? (EN_BUILTIN[format] ?? '') : format, EN_CODES)?.[0]?.toks ?? []
  let time = false
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i]
    if (t.t === 'elapsed' || t.t === 'ampm') time = true
    else if (t.t === 'date') {
      if (t.s[0] === 'h' || t.s[0] === 's' || (t.s[0] === 'm' && minutes(toks, i))) time = true
      else return 'date'
    }
  }
  return time ? 'time' : 'number'
}

/** A date's serial number in the book's system, or null for one out of it. */
function dateSerial(text: string, book: BookInput): number | null {
  try {
    return dateToSerial(text, { dateSystem: book.dateSystem })
  } catch {
    return null
  }
}

/** The case a text is written in, put on another: all capitals, all small, or a capital first. */
function casing(s: string): (x: string) => string {
  if (s === s.toUpperCase() && s !== s.toLowerCase()) return (x) => x.toUpperCase()
  if (s === s.toLowerCase()) return (x) => x.toLowerCase()
  return (x) => x[0].toUpperCase() + x.slice(1).toLowerCase()
}

/** The runs of a line's source: consecutive cells of one kind (text with one prefix, names of one list). */
function runsOf(items: Item[]): Run[] {
  const runs: Run[] = []
  let run: Run | null = null
  items.forEach((x, at) => {
    const last = run?.items.at(-1)
    if (run && last && together(last, x)) run.items.push(x)
    else runs.push((run = { at, items: [x] }))
  })
  return runs
}

const together = (x: Item, y: Item): boolean => x.k !== 'other' && x.k === y.k && (x.k !== 'text' || x.key === (y as typeof x).key)

/**
 * A run's value at element `e` (0 its first cell, its length the first copied), given the source cell it copies; null when
 * the run is copied. A single number goes on by 1 when `single`, else it is copied.
 */
function goesOn(run: Run, single: boolean, epoch: number): ((e: number, source: Item) => CellInput['value']) | null {
  const first = run.items[0]
  const ys = run.items.map((x) => (x.k === 'other' ? 0 : x.y))
  const m = ys.length
  switch (first.k) {
    case 'other':
      return null
    case 'number': {
      if (m === 1) return single ? (e) => ys[0] + e : null
      const f = trend(ys)
      return (e) => f(e)
    }
    case 'time': {
      if (m === 1) return (e) => ys[0] + e / 24
      return trend(ys)
    }
    case 'date': {
      if (m === 1) return (e) => ys[0] + e
      const months = monthStep(ys, epoch)
      return months ?? trend(ys)
    }
    case 'text': {
      const f = m === 1 ? (e: number) => ys[0] + e : trend(ys)
      return (e, source) => (source.k === 'text' ? source : first).make(Math.round(f(e)))
    }
  }
}

/** The values' linear trend at element `e`: their step when they have one, else the least-squares line. */
function trend(ys: number[]): (e: number) => number {
  const d = ys[1] - ys[0]
  if (ys.every((y, i) => i === 0 || y - ys[i - 1] === d)) return (e) => ys[0] + d * e
  const m = ys.length
  const mx = (m - 1) / 2
  const my = ys.reduce((s, y) => s + y, 0) / m
  let sxy = 0
  let sxx = 0
  ys.forEach((y, x) => {
    sxy += (x - mx) * (y - my)
    sxx += (x - mx) * (x - mx)
  })
  const slope = sxy / sxx
  return (e) => my + slope * (e - mx)
}

/**
 * Dates a whole number of months apart, each step the same, a day past a month's end taken back to its last day (Jan 31,
 * Feb 28, Mar 31): the date `e` steps on from the first, else null.
 */
function monthStep(ys: number[], epoch: number): ((e: number) => number) | null {
  const { y, m, d } = ymd(Math.floor(ys[0]) + epoch)
  const t = ys[0] - Math.floor(ys[0])
  const second = ymd(Math.floor(ys[1]) + epoch)
  const step = second.y * 12 + second.m - (y * 12 + m)
  if (step === 0) return null
  const at = (e: number) => {
    const to = addMonths(y, m, step * e)
    return serialOf(to.y, to.m, Math.min(d, daysInMonth(to.y, to.m))) - epoch + t
  }
  return ys.every((v, i) => Math.abs(at(i) - v) < 1e-9) ? at : null
}
