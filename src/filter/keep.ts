/**
 * Which cells a column's filter keeps, as Excel applies the criteria a workbook holds (Data > Reapply over a file's
 * `<filterColumn>`): values by the text a cell shows, dates by their parts, custom conditions by the operand read the
 * same way in every locale.
 */
import type { CellStyle, CellValue, Color, ColorFilter, ColumnFilter, ConditionalAppearance, CustomFilter, DateGroup, DynamicFilter, ErrorCode, FilterOperator, IconFilter, NumberFormat, TextRun, TopFilter, ValuesFilter } from '../contract.js'
import { addMonths, critTextToNumber, enCritNumber, momentOf, serialExcel, trimSpaces, ymd } from '../dates.js'
import { shownAs } from '../numfmt/shown-as.js'
import { ERROR_CODE } from '../syntax/errors.js'
import { Wildcard } from '../wildcards.js'

/**
 * A column's cells, by their place below the header: the value, the text it shows in no column's room, the number
 * format, the style, the runs of a rich text and the appearance its conditional formats give it.
 */
export interface ColumnCells {
  n: number
  value(i: number): CellValue
  text(i: number): string
  format(i: number): NumberFormat | undefined
  style(i: number): CellStyle | undefined
  runs(i: number): TextRun[] | undefined
  look(i: number): ConditionalAppearance | undefined
}

/**
 * What a filter compares by: the order of two texts in the workbook's locale, its date system's epoch (`DateSystem`),
 * and the day of the calculation, a serial number of that system.
 */
export interface KeepRules {
  compareText(a: string, b: string): number
  epoch: number
  today: number
}

/** The cells a filter keeps and the filter with the bounds it found; or the first cell of an error a top or average filter cannot be applied over. */
export type Kept = { keep: (i: number) => boolean; filter?: ColumnFilter } | { error: number }

/** A cell that holds an error. */
const isError = (v: CellValue): v is { error: ErrorCode } => typeof v === 'object' && v !== null

/** Text Excel's filter takes as the same in any case: `Ё` is `ё` but not `е`, `ß` is `ss` (the filter sets `values-30`, `values-31`). */
const textKey = (s: string) => s.normalize('NFD').toLowerCase().replace(/ß/g, 'ss')

/** A cell a filter takes as blank: none, an empty text (a formula's `""`) or spaces alone; a no-break space is not. */
const isBlank = (v: CellValue) => v === null || (typeof v === 'string' && trimSpaces(v) === '')

/** The parts of a date chosen, down to its level, with Excel's defaults for those it lacks. */
function dateParts(g: DateGroup): number[] {
  const all = [g.year, g.month ?? 1, g.day ?? 1, g.hour ?? 0, g.minute ?? 0, g.second ?? 0]
  return all.slice(0, ['year', 'month', 'day', 'hour', 'minute', 'second'].indexOf(g.level) + 1)
}

/** The year, month and day of a day in the calendar of a filter's dates (Excel's Hijri is the tabular one); Gregorian for the others. */
function calendarOf(calendar: string | undefined): (y: number, m: number, d: number) => [number, number, number] {
  if (calendar !== 'hijri') return (y, m, d) => [y, m, d]
  const f = new Intl.DateTimeFormat('en-u-ca-islamic-tbla', { timeZone: 'UTC', year: 'numeric', month: 'numeric', day: 'numeric' })
  return (y, m, d) => {
    const t = new Date(0)
    t.setUTCFullYear(y, m - 1, d)
    const parts = Object.fromEntries(f.formatToParts(t).map((p) => [p.type, p.value]))
    return [Number(parts.year), Number(parts.month), Number(parts.day)]
  }
}

/** A values filter: the values chosen by the text a cell shows, its spaces at either end left out, blanks, and dates by their parts. */
function valuesKeep(f: ValuesFilter, cells: ColumnCells, rules: KeepRules): Kept {
  const chosen = new Set((f.values ?? []).map(textKey))
  const dates = (f.dates ?? []).map(dateParts)
  const calendar = dates.length ? calendarOf(f.calendar) : null
  const keep = (i: number) => {
    const v = cells.value(i)
    if (isBlank(v)) return f.blank === true
    if (chosen.size && chosen.has(textKey(trimSpaces(cells.text(i))))) return true
    // a number shown as a date: a time of day alone, or a serial in General, is no date
    if (!calendar || typeof v !== 'number' || shownAs(cells.format(i)) !== 'date') return false
    const { day, second } = momentOf(v + rules.epoch)
    const { y, m, d } = ymd(day)
    const parts = [...calendar(y, m, d), Math.floor(second / 3600), Math.floor(second / 60) % 60, second % 60]
    return dates.some((g) => g.every((x, k) => x === parts[k]))
  }
  return { keep }
}

/** A custom condition's operand as the file's reader takes it, in every locale: a number, percent or date in en-US forms, TRUE or FALSE, an error, or text. */
type Operand = { kind: 'number'; n: number } | { kind: 'bool'; b: boolean } | { kind: 'error'; code: ErrorCode } | { kind: 'text'; key: string; text: string; wild: Wildcard | null } | { kind: 'empty' } | { kind: 'blank' }

/**
 * The operand of a condition: an empty one apart, spaces alone are a blank, which Excel writes for its (Blanks) and
 * (NonBlanks) conditions (`notEqual " "`).
 */
function operandOf(text: string, epoch: number): Operand {
  if (text === '') return { kind: 'empty' }
  if (trimSpaces(text) === '') return { kind: 'blank' }
  const n = enCritNumber(text)
  if (n !== null) return { kind: 'number', n }
  const date = critTextToNumber(text)
  if (date !== null) return { kind: 'number', n: date - epoch }
  if (/^(TRUE|FALSE)$/i.test(text)) return { kind: 'bool', b: /^TRUE$/i.test(text) }
  const code = ERROR_CODE[text.toUpperCase()]
  if (code) return { kind: 'error', code }
  return { kind: 'text', key: textKey(text), text, wild: /[*?~]/.test(text) ? new Wildcard(text, 'i') : null }
}

/**
 * One condition, as Excel reapplies it (the sets `parse-*`, `extra-custom-*`): equality by the text a cell shows, its
 * spaces at either end left out; the other operators compare a value with an operand of its kind, numbers by value and
 * text by the locale's order, FALSE before TRUE, a pattern of wildcards as its own table says.
 */
function conditionKeep(operator: FilterOperator, text: string, cells: ColumnCells, rules: KeepRules): (i: number) => boolean {
  const o = operandOf(text, rules.epoch)
  const shown = (i: number) => trimSpaces(cells.text(i))
  if (o.kind === 'blank' && (operator === 'equal' || operator === 'notEqual')) return operator === 'equal' ? (i) => isBlank(cells.value(i)) : (i) => !isBlank(cells.value(i))
  if (operator === 'equal') {
    if (o.kind === 'empty') return (i) => isBlank(cells.value(i))
    const key = textKey(text)
    const wild = /[*?~]/.test(text) ? new Wildcard(text, 'i') : null
    return (i) => (isBlank(cells.value(i)) ? false : wild ? wild.test(shown(i)) : textKey(shown(i)) === key)
  }
  const order = (i: number): number | null => {
    const v = cells.value(i)
    if (o.kind === 'bool') return typeof v === 'boolean' ? Number(v) - Number(o.b) : null
    if (o.kind === 'number') return typeof v === 'number' ? Math.sign(v - o.n) : null
    if (o.kind === 'error') return typeof v === 'object' && v !== null && v.error === o.code ? 0 : null
    // an empty text, a formula's `""` or spaces, comes before any other
    if (o.kind !== 'text' || typeof v !== 'string') return null
    return textKey(trimSpaces(v)) === o.key ? 0 : Math.sign(rules.compareText(trimSpaces(v), o.text))
  }
  if (operator === 'notEqual') {
    if (o.kind === 'empty') return (i) => isBlank(cells.value(i))
    return (i) => {
      const v = cells.value(i)
      if (o.kind === 'text' && o.wild && typeof v === 'string' && !isBlank(v)) return !o.wild.test(trimSpaces(v))
      return order(i) !== 0
    }
  }
  if (o.kind === 'text' && o.wild) {
    // a pattern orders text as Excel's table of `a*` says: at most what matches it, more what does not, an empty text
    // too, any text at least
    const wild = o.wild
    const text = (i: number) => {
      const v = cells.value(i)
      return typeof v === 'string' ? trimSpaces(v) : null
    }
    switch (operator) {
      case 'lessThan':
        return () => false
      case 'lessThanOrEqual':
        return (i) => text(i) !== null && wild.test(text(i)!)
      case 'greaterThan':
        return (i) => text(i) !== null && !wild.test(text(i)!)
      case 'greaterThanOrEqual':
        return (i) => text(i) !== null
    }
  }
  const pass = { lessThan: (x: number) => x < 0, lessThanOrEqual: (x: number) => x <= 0, greaterThan: (x: number) => x > 0, greaterThanOrEqual: (x: number) => x >= 0 }[operator]
  return (i) => {
    const x = order(i)
    return x !== null && pass(x)
  }
}

/** A custom filter: its one or two conditions, both or either. */
function customKeep(f: CustomFilter, cells: ColumnCells, rules: KeepRules): Kept {
  const tests = f.conditions.map((c) => conditionKeep(c.operator ?? 'equal', c.value, cells, rules))
  return { keep: f.and ? (i) => tests.every((t) => t(i)) : (i) => tests.some((t) => t(i)) }
}

/** The numbers of a column, dates too, without text that looks like one, booleans and blanks; the first error stops it. */
function numbersOf(cells: ColumnCells): number[] | { error: number } {
  const out: number[] = []
  for (let i = 0; i < cells.n; i++) {
    const v = cells.value(i)
    if (isError(v)) return { error: i }
    if (typeof v === 'number') out.push(v)
  }
  return out
}

/**
 * The top or bottom items, or percent of the column's numbers, whatever the other columns keep (the filter sets
 * `follow-rank-*`): the `k`-th from the end counts, ties with it too; a percent of `n` numbers is `n * percent / 100`
 * rounded down, at least 1. `filterValue` is that number.
 */
function topKeep(f: TopFilter, cells: ColumnCells): Kept {
  const nums = numbersOf(cells)
  if (!Array.isArray(nums)) return nums
  if (!nums.length) return { keep: () => false }
  const n = nums.length
  const k = Math.max(1, Math.min(n, Math.floor(f.percent ? (n * f.value) / 100 : f.value)))
  nums.sort((a, b) => (f.bottom ? a - b : b - a))
  const bound = nums[k - 1]
  const keep = (i: number) => {
    const v = cells.value(i)
    return typeof v === 'number' && (f.bottom ? v <= bound : v >= bound)
  }
  return { keep, filter: { ...f, filterValue: bound } }
}

/** A serial number of the workbook's date system as Excel's ISO text of it, as a file's `valIso`. */
function isoOf(serial: number, epoch: number): string {
  const { day, second } = momentOf(serial + epoch)
  const { y, m, d } = ymd(day)
  const two = (x: number) => String(x).padStart(2, '0')
  return `${String(y).padStart(4, '0')}-${two(m)}-${two(d)}T${two(Math.floor(second / 3600))}:${two(Math.floor(second / 60) % 60)}:${two(second % 60)}`
}

/**
 * The days a relative date filter keeps, from the first to before the last, as serial numbers of the workbook's date
 * system: weeks from Sunday, the year to date from January 1 (the filter sets `file-dynamic-*`).
 */
function period(kind: DynamicFilter['kind'], today: number, epoch: number): [number, number] | null {
  const day = (y: number, m: number) => serialExcel(y, m, 1) - epoch
  const { y, m } = ymd(today + epoch)
  const month = (from: number, months: number): [number, number] => {
    const a = addMonths(y, from, 0)
    const b = addMonths(y, from, months)
    return [day(a.y, a.m), day(b.y, b.m)]
  }
  const sunday = today - ((today + epoch - 1) % 7)
  const q = m - ((m - 1) % 3)
  switch (kind) {
    case 'yesterday':
      return [today - 1, today]
    case 'today':
      return [today, today + 1]
    case 'tomorrow':
      return [today + 1, today + 2]
    case 'lastWeek':
      return [sunday - 7, sunday]
    case 'thisWeek':
      return [sunday, sunday + 7]
    case 'nextWeek':
      return [sunday + 7, sunday + 14]
    case 'lastMonth':
      return month(m - 1, 1)
    case 'thisMonth':
      return month(m, 1)
    case 'nextMonth':
      return month(m + 1, 1)
    case 'lastQuarter':
      return month(q - 3, 3)
    case 'thisQuarter':
      return month(q, 3)
    case 'nextQuarter':
      return month(q + 3, 3)
    case 'lastYear':
      return [day(y - 1, 1), day(y, 1)]
    case 'thisYear':
      return [day(y, 1), day(y + 1, 1)]
    case 'nextYear':
      return [day(y + 1, 1), day(y + 2, 1)]
    case 'yearToDate':
      return [day(y, 1), today + 1]
    default:
      return null
  }
}

/**
 * A dynamic filter: above or below the average of the column's numbers, strictly, whatever the other columns keep; or
 * the dates of a period, by a number with a format of a date: relative to the day of the calculation, with the period's
 * bounds stored, or a month or quarter of any year.
 */
function dynamicKeep(f: DynamicFilter, cells: ColumnCells, rules: KeepRules): Kept {
  if (f.kind === 'null') return { keep: () => true }
  if (f.kind === 'aboveAverage' || f.kind === 'belowAverage') {
    const nums = numbersOf(cells)
    if (!Array.isArray(nums)) return nums
    if (!nums.length) return { keep: () => false }
    const average = nums.reduce((a, b) => a + b, 0) / nums.length
    const above = f.kind === 'aboveAverage'
    const keep = (i: number) => {
      const v = cells.value(i)
      return typeof v === 'number' && (above ? v > average : v < average)
    }
    return { keep, filter: { ...f, value: average } }
  }
  const date = (i: number) => {
    const v = cells.value(i)
    return typeof v === 'number' && shownAs(cells.format(i)) === 'date' ? v : null
  }
  const bounds = period(f.kind, rules.today, rules.epoch)
  if (bounds) {
    const [from, to] = bounds
    const keep = (i: number) => {
      const v = date(i)
      return v !== null && v >= from && v < to
    }
    const iso = f.valueIso !== undefined || f.maxValueIso !== undefined ? { valueIso: isoOf(from, rules.epoch), maxValueIso: isoOf(to, rules.epoch) } : {}
    return { keep, filter: { ...f, value: from, maxValue: to, ...iso } }
  }
  // Q1 to Q4 and M1 to M12: the months of any year
  const months = f.kind[0] === 'Q' ? [1, 2, 3].map((x) => (Number(f.kind[1]) - 1) * 3 + x) : [Number(f.kind.slice(1))]
  return {
    keep: (i) => {
      const v = date(i)
      return v !== null && months.includes(ymd(momentOf(v + rules.epoch).day).m)
    },
  }
}

/** White, which a file's automatic color reads as: a filter by the automatic font color keeps a cell whose font has no color of its own. */
const WHITE = '#FFFFFF'

/**
 * A font's color as a filter compares it: black is the automatic color, which a reader gives a rich text's runs and
 * leaves out of a cell's style; a filter by black keeps neither (the sets `reload-font-black`, `color-9-0`).
 */
const fontColor = (c: Color | undefined) => (c === '#000000' ? undefined : c)

/**
 * A color filter: by the fill a conditional format gives a cell, else its own; or by its font's color, a conditional
 * format's, else its own or a run's of its rich text, any one of them (the filter sets `color-8-*`, `color-9-*`). A
 * filter without a color keeps the cells without a fill, or with the automatic font color.
 */
function colorKeep(f: ColorFilter, cells: ColumnCells): Kept {
  const want: Color | undefined = f.font ? (f.style.font?.color ?? f.style.fill?.color) : (f.style.fill?.color ?? f.style.font?.color)
  if (!f.font) return { keep: (i) => (cells.look(i)?.style?.fill?.color ?? cells.style(i)?.fill?.color) === want }
  const same = (c: Color | undefined) => fontColor(c) === (want === WHITE ? undefined : want)
  return {
    keep: (i) => {
      const cf = cells.look(i)?.style?.font?.color
      if (cf !== undefined) return same(cf)
      const own = cells.style(i)?.font?.color
      const runs = typeof cells.value(i) === 'string' ? cells.runs(i) : undefined
      return runs?.length ? (runs[0].start > 0 && same(own)) || runs.some((r) => same(r.font ? r.font.color : own)) : same(own)
    },
  }
}

/** An icon filter: the icon a conditional format's icon set gives a cell, by its set and place; without one, the cells without an icon. */
function iconKeep(f: IconFilter, cells: ColumnCells): Kept {
  return {
    keep: (i) => {
      const icon = cells.look(i)?.icon
      return f.icon === undefined ? !icon : !!icon && icon.set === f.set && icon.index === f.icon
    },
  }
}

/** How each kind of filter keeps its cells. */
export const KEEPS: { [K in ColumnFilter['type']]: (f: Extract<ColumnFilter, { type: K }>, cells: ColumnCells, rules: KeepRules) => Kept } = {
  values: valuesKeep,
  custom: customKeep,
  top: topKeep,
  dynamic: dynamicKeep,
  color: colorKeep,
  icon: iconKeep,
}
