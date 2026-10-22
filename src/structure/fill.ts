/**
 * Excel's fill handle (`StructureOp` `fillRange`, `Range.AutoFill`), as measured on Excel for Windows (the struct sets
 * `fill`, `fill-held`, `fill-third-held`). The cells added take the source's cells in turn as copies of it: whole sources
 * next to it, then the part of it nearest to it at the far end (`copyRange` each, so formulas, looks, rules and
 * validations go as a copy takes them). Then, unless the fill copies alone, each line along the fill (a column of a fill
 * down or up, a row of one right or left) is cut into runs of its source's cells that go on together, and the copies of a
 * run's cells take the run's next values. A run is cells of one kind in a row, empty cells between them skipped and text,
 * booleans and formulas ending it: numbers go on by their trend, dates by their step in days or by months, times by their
 * step or by hours, text with a number by the number, names of days and months through their list; dates and text with
 * no single step are copied. A run's cells come again in every copy of the source, so the run goes on from copy to copy
 * (1, 2, x → 3, 4, x, 5, 6, x), and backwards in a fill up or left.
 */
import type { BookInput, CellInput, FillType } from '../contract.js'
import { addMonths, daysInMonth, epochOf, serialOf, ymd } from '../dates.js'
import { dateToSerial, isDateSerial, serialToDate } from '../date-serial.js'
import { InputError } from '../input-error.js'
import { entryAt, lines as linesOf } from '../layout/spans.js'
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
import { checkFillWrites } from './writes.js'
import { renamedByPaste } from '../tables/moves.js'
import { tablesOf } from '../tables/model.js'

/** A copy of part of the sheet over another part, as `editStructure` applies one (`copyRange`). */
export type Copier = (book: BookInput, from: Rect, to: Rect) => Applied

/** A cell of a line's source as a fill reads it: what kind of series it can go on in, and its value there. */
type Item =
  | { k: 'number' | 'date' | 'time'; y: number }
  /**
   * Text that goes on by a number in it: the cells of one `key` go on together, `make` writes the text of another number;
   * a list's names and quarters go round `cycle`.
   */
  | { k: 'text'; key: string; y: number; make: (n: number) => string; cycle?: number }
  | { k: 'empty' }
  | { k: 'other' }

type Kind = Exclude<Item, { k: 'empty' | 'other' }>

/** A run of a line's source: cells that go on together, at their positions along the source. */
interface Run {
  at: number[]
  items: Kind[]
}

/**
 * The fill of `src` over `target` on `sheet` (which holds `src` and reaches past it along one axis), its copies made by
 * `copy`; `to` is the target as given, for the message of one of the wrong shape.
 */
export function fillRange(book: BookInput, sheet: string, src: Rect, target: Rect, to: string, type: FillType, copy: Copier): Applied {
  const way = wayOf(src, target)
  if (!way) throw new InputError('op.to', `expected a range that contains op.range and extends past it down, up, right or left, got ${JSON.stringify(to)}`)
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
  // the copies count together against limits.writes: each checks its own cells too, which is then within the limit
  checkFillWrites(book, sheetIndex(book, sheet), pieces)
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
  const added: [number, number] = forward ? [b + 1, b + n] : [a - n, a - 1]
  const written = seriesCells(book, s, src, axis, added, type)
  if (!written.length) return { ok: true, book: out, geometry }
  const cells = { ...out.sheets[s].cells }
  for (const [addr, value] of written) cells[addr] = { ...cells[addr], value }
  const filled = { ...out, sheets: out.sheets.map((x, i) => (i === s ? { ...x, cells } : x)) }
  // a series over a table's header cells renames their columns to what it wrote, as a paste over them does
  return { ok: true, book: renamedByPaste(tablesOf(filled), filled, [[s, withBounds(src, axis, ...added)]], () => true), geometry }
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

/** A cell that holds nothing a fill reads: none, or one without a value or a formula. */
const empty = (cell: CellInput | undefined) => !cell || (cell.formula === undefined && (cell.value === undefined || cell.value === null))

/** The cells of positions `added` (from, to along the axis) that a series writes, with their values. */
function seriesCells(book: BookInput, s: number, src: Rect, axis: Axis, added: [number, number], type: FillType): [string, CellInput['value']][] {
  const sheet = book.sheets[s]
  const g = gridOf(sheet)
  // a cell's number format: its own, else its row's, else its column's, the sheet's rows and columns read once
  const [rows, cols] = [linesOf(sheet.rows, 'r'), linesOf(sheet.columns, 'c')]
  const formatAt = (cell: CellInput | undefined, r: number, c: number) => cell?.numberFormat ?? entryAt(rows, r)?.numberFormat ?? entryAt(cols, c)?.numberFormat
  const a = lo(src, axis)
  const h = hi(src, axis) - a + 1
  // the source's cells by line across the fill, at their position along it
  const lines = new Map<number, Map<number, CellInput>>()
  // a number alone in its run is copied, as a single cell is, when the source holds numbers alone and the fill is no series
  let numbers = type !== 'series'
  for (const i of cellsIn(g, src)) {
    const cell = g.cell[i]
    if (empty(cell)) continue
    numbers &&= cell.formula === undefined && typeof cell.value === 'number' && shownAs(formatAt(cell, g.row[i], g.col[i])) === 'number'
    const [along, across] = axis === 'r' ? [g.row[i], g.col[i]] : [g.col[i], g.row[i]]
    let line = lines.get(across)
    if (!line) lines.set(across, (line = new Map()))
    line.set(along - a, cell)
  }
  const lists = listsOf(book)
  const epoch = epochOf(book.dateSystem)
  const out: [string, CellInput['value']][] = []
  for (const [across, line] of lines) {
    const items = Array.from({ length: h }, (_, p) => {
      const [r, c] = axis === 'r' ? [a + p, across] : [across, a + p]
      const cell = line.get(p)
      return itemOf(cell, formatAt(cell, r, c), book, lists)
    })
    for (const run of runsOf(items)) {
      const next = goesOn(run, !numbers, epoch)
      if (!next) continue
      const m = run.at.length
      for (let pos = added[0]; pos <= added[1]; pos++) {
        const p = pos - a
        const d = run.at.indexOf(((p % h) + h) % h)
        if (d < 0) continue
        const e = Math.floor(p / h) * m + d
        const [r, c] = axis === 'r' ? [pos, across] : [across, pos]
        const value = next(e, run.items[d])
        const own = line.get(run.at[d])!.value
        // a date past the date system is written as its serial number, as a fill over numbers in a date format writes it
        const date = typeof own === 'object' && own !== null && typeof value === 'number' && isDateSerial(value, epoch)
        out.push([formatAddress(r, c), date ? { date: serialToDate(value, { dateSystem: book.dateSystem }) } : value])
      }
    }
  }
  return out
}

/** The lists a text goes on through: en-US's short and long day and month names, and the book's locale's. */
function listsOf(book: BookInput): (readonly string[])[] {
  const en = [EN_NAMES.daysShort, EN_NAMES.days, EN_NAMES.monthsShort, EN_NAMES.months]
  const r = localeOf(book.locale).regional
  return r && r.lcid !== '409' ? [...en, r.daysShort, r.days, r.monthsShort, r.months] : en
}

/** A quarter: `Q1`, `Qtr 2`, `Quarter 4`. */
const QUARTER = /^(q|qtr|quarter)( ?)([1-4])$/i
/** An ordinal of en-US: `1st`, `22nd`. */
const ORDINAL = /^(\d+)(st|nd|rd|th)$/i
/** Text that starts with a number and a space: its digits and the text from the space on. */
const LEADING = /^(\d+)( .*)$/s
/** Text that ends in a number: the text before it and its digits. */
const SUFFIX = /^(.*?)(\d+)$/s

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
  if (empty(cell)) return { k: 'empty' }
  if (cell!.formula !== undefined) return { k: 'other' }
  const v = cell!.value
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
    if (y >= 0) return { k: 'text', key: `list ${i}`, y, make: (n) => caseOf(list[round(n, list.length)]), cycle: list.length }
  }
  let m = QUARTER.exec(v)
  if (m) {
    const head = m[1] + m[2]
    return { k: 'text', key: `quarter ${head.toLowerCase()}`, y: Number(m[3]), make: (n) => head + (round(n - 1, 4) + 1), cycle: 4 }
  }
  if ((m = ORDINAL.exec(v))) return { k: 'text', key: 'ordinal', y: Number(m[1]), make: (n) => Math.abs(n) + ordinal(n) }
  if ((m = LEADING.exec(v))) {
    const tail = m[2]
    return { k: 'text', key: `leading ${tail}`, y: Number(m[1]), make: (n) => digits(n, 0) + tail }
  }
  if ((m = SUFFIX.exec(v))) {
    const [head, width] = [m[1], widthOf(m[2])]
    return { k: 'text', key: `suffix ${head}`, y: Number(m[2]), make: (n) => head + digits(n, width) }
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

/** The runs of a line's source: cells of one kind (text with one key), empty cells between them skipped, other cells ending them. */
function runsOf(items: Item[]): Run[] {
  const runs: Run[] = []
  let run: Run | null = null
  items.forEach((x, at) => {
    if (x.k === 'empty') return
    if (x.k === 'other') return void (run = null)
    const last = run?.items.at(-1)
    if (run && last && last.k === x.k && (x.k !== 'text' || (last as typeof x).key === x.key)) {
      run.at.push(at)
      run.items.push(x)
    } else runs.push((run = { at: [at], items: [x] }))
  })
  return runs
}

/**
 * A run's value at element `e` (0 its first cell, its length the first copied), given the source cell it copies; null when
 * the run is copied. A single number goes on by 1 when `single`, else it is copied.
 */
function goesOn(run: Run, single: boolean, epoch: number): ((e: number, source: Kind) => CellInput['value']) | null {
  const first = run.items[0]
  const ys = run.items.map((x) => x.y)
  if (ys.length === 1 && first.k === 'number') return single ? (e) => ys[0] + e : null
  if (ys.length === 1 && first.k === 'time') return (e) => ys[0] + e / 24
  switch (first.k) {
    case 'number':
      return trend(ys)
    case 'time':
      return stepOf(ys)
    case 'date':
      return (ys.length > 1 && monthStep(ys, epoch)) || stepOf(ys)
    case 'text': {
      const f = stepOf(first.cycle ? unwrapped(ys, first.cycle) : ys)
      return f && ((e, source) => (source.k === 'text' ? source : first).make(Math.round(f(e))))
    }
  }
}

/** Positions in a list that goes round `cycle`, each taken a step forward from the one before it (Sat, Sun as 6, 7). */
const unwrapped = (ys: number[], cycle: number) => ys.map((y, i) => (i === 0 ? y : ys[0] + ys.slice(1, i + 1).reduce((s, x, j) => s + round(x - ys[j], cycle), 0)))

/** Two values alike in Excel's 15 digits (9:00 and two hours on from 9:00 part in the last bit from 11:00). */
const same = (x: number, y: number) => x.toPrecision(15) === y.toPrecision(15)

/** The values' step, when each goes on from the one before it by the same (a single value by 1): the value at element `e`; else null. */
function stepOf(ys: number[]): ((e: number) => number) | null {
  const d = ys.length > 1 ? ys[1] - ys[0] : 1
  return ys.every((y, i) => same(ys[0] + d * i, y)) ? (e) => ys[0] + d * e : null
}

/** The values' linear trend at element `e`: their step when they have one, else the least-squares line. */
function trend(ys: number[]): (e: number) => number {
  const step = stepOf(ys)
  if (step) return step
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
 * Dates on one day of the month, or each on its month's last day, a whole number of months apart, each step the same: the
 * date `e` steps on from the first (on the last day of its month, or on the day, taken back to the last day of a shorter
 * month); else null.
 */
function monthStep(ys: number[], epoch: number): ((e: number) => number) | null {
  const ds = ys.map((y) => ymd(Math.floor(y) + epoch))
  const ends = ds.every((x) => x.d === daysInMonth(x.y, x.m))
  if (!ends && ds.some((x) => x.d !== ds[0].d)) return null
  const months = ds.map((x) => x.y * 12 + x.m)
  const step = months[1] - months[0]
  if (step === 0 || months.some((x, i) => i > 0 && x - months[i - 1] !== step)) return null
  const { y, m, d } = ds[0]
  const t = ys[0] - Math.floor(ys[0])
  return (e) => {
    const to = addMonths(y, m, step * e)
    const last = daysInMonth(to.y, to.m)
    return serialOf(to.y, to.m, ends ? last : Math.min(d, last)) - epoch + t
  }
}
