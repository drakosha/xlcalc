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
  | { k: 'suffix'; prefix: string; y: number; width: number }
  | { k: 'list'; list: readonly string[]; y: number; caseOf: (s: string) => string }
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
    for (const run of runsOf(items)) {
      const next = goesOn(run, type, epoch)
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
  const key = v.toLowerCase()
  for (const list of lists) {
    const y = list.findIndex((x) => x.toLowerCase() === key)
    if (y >= 0) return { k: 'list', list, y, caseOf: casing(v) }
  }
  const m = SUFFIX.exec(v)
  if (m) return { k: 'suffix', prefix: m[1], y: Number(m[2]), width: m[2].length > 1 && m[2][0] === '0' ? m[2].length : 0 }
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

const together = (x: Item, y: Item): boolean =>
  x.k !== 'other' && x.k === y.k && (x.k !== 'suffix' || x.prefix === (y as typeof x).prefix) && (x.k !== 'list' || x.list === (y as typeof x).list)

/** A run's value at element `e` (0 its first cell, its length the first copied), given the source cell it copies; null when the run is copied. */
function goesOn(run: Run, type: FillType, epoch: number): ((e: number, source: Item) => CellInput['value']) | null {
  const first = run.items[0]
  const ys = run.items.map((x) => (x.k === 'other' ? 0 : x.y))
  const m = ys.length
  switch (first.k) {
    case 'other':
      return null
    case 'number': {
      if (m === 1) return type === 'series' ? (e) => ys[0] + e : null
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
    case 'suffix': {
      const f = m === 1 ? (e: number) => ys[0] + e : trend(ys)
      return (e) => {
        const v = Math.abs(Math.round(f(e)))
        return first.prefix + String(v).padStart(first.width, '0')
      }
    }
    case 'list': {
      const n = first.list.length
      const f = m === 1 ? (e: number) => ys[0] + e : trend(ys)
      return (e, source) => {
        const i = ((Math.round(f(e)) % n) + n) % n
        return (source.k === 'list' ? source.caseOf : first.caseOf)(first.list[i])
      }
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

/** Dates on one day of the month a whole number of months apart, each step the same: the date `e` steps on, else null. */
function monthStep(ys: number[], epoch: number): ((e: number) => number) | null {
  const ds = ys.map((y) => ({ ...ymd(Math.floor(y) + epoch), t: y - Math.floor(y) }))
  if (ds.some((x) => x.d !== ds[0].d)) return null
  const step = ds[1].y * 12 + ds[1].m - (ds[0].y * 12 + ds[0].m)
  if (step === 0 || ds.some((x, i) => i > 0 && x.y * 12 + x.m - (ds[i - 1].y * 12 + ds[i - 1].m) !== step)) return null
  const { y, m, d, t } = ds[0]
  return (e) => {
    const at = addMonths(y, m, step * e)
    return serialOf(at.y, at.m, Math.min(d, daysInMonth(at.y, at.m))) - epoch + t
  }
}
