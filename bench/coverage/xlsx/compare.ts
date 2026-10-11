/**
 * Comparing a book with the one read back from a file, and Excel's stored values with the engine's. Both sides
 * are flattened to entries by place (`cell S!A1`, `column S!B`, `name S!rate` …) holding plain values: a style by
 * its fields rather than its index, a formula as its canonical text, a date as its serial number in the book's date
 * system, a code equal to a built-in format every locale shows alike as that format's number, spans of rows and
 * columns line by line (a line of size 0 hidden), iterative calculation with Excel's defaults filled in.
 */
import type { BookInput, CellInput, CellResult, CellStyle, CellValue, ComputeResult, NumberFormat, TableInput } from '../../../src/contract.js'
import { entryAt, lines } from '../../../src/layout/spans.js'
import { epochOf } from '../../../src/dates.js'
import { dateToSerial, indexToCol, parseAddress } from '../../../src/model.js'
import { canonical, parse, validate } from '../../../src/syntax.js'
import { tableAt, tablesOf, type Tables } from '../../../src/tables/model.js'
import { BUILTIN_ID } from '../../../src/xlsx/styles/numfmt.js'
import { numberFormatOf } from '../stored.js'
import type { XlsxCached } from '../../../src/xlsx/types.js'

export interface Difference {
  place: string
  written?: unknown
  read?: unknown
}

const DEFAULT_FONT = { name: 'Calibri', size: 11 }

/** A style as the file keeps it: font fields equal to the default font's left out, colours upper-case, nothing empty. */
function normStyle(s: CellStyle | undefined): CellStyle | undefined {
  if (!s) return undefined
  const up = <T extends { color?: string }>(x: T): T => (x.color ? { ...x, color: x.color.toUpperCase() } : x)
  const out: CellStyle = {}
  if (s.font) {
    const font = up({ ...s.font })
    if (font.name === DEFAULT_FONT.name) delete font.name
    if (font.size === DEFAULT_FONT.size) delete font.size
    for (const k of ['bold', 'italic', 'strike'] as const) if (font[k] === false) delete font[k]
    if (Object.keys(font).length) out.font = font
  }
  if (s.fill) out.fill = up(s.fill)
  if (s.border && Object.keys(s.border).length) out.border = Object.fromEntries(Object.entries(s.border).map(([k, l]) => [k, up(l)]))
  if (s.alignment) {
    const a = { ...s.alignment }
    if (a.horizontal === 'general') delete a.horizontal
    if (a.vertical === 'bottom') delete a.vertical
    if (a.wrap === false) delete a.wrap
    if (a.indent === 0) delete a.indent
    if (Object.keys(a).length) out.alignment = a
  }
  return Object.keys(out).length ? out : undefined
}

/** JSON with sorted keys, so that equal objects give equal text. */
function stable(x: unknown): string {
  if (x === undefined) return 'undefined'
  if (x === null || typeof x !== 'object') return JSON.stringify(x)
  if (Array.isArray(x)) return `[${x.map(stable).join(',')}]`
  return `{${Object.keys(x)
    .filter((k) => (x as Record<string, unknown>)[k] !== undefined)
    .sort()
    .map((k) => `${JSON.stringify(k)}:${stable((x as Record<string, unknown>)[k])}`)
    .join(',')}}`
}

/**
 * A formula's canonical text; a name's formula may be a union without parentheses. With the book's tables its structured
 * references are bound to them (`[@Qty]` in a table's cell is `Sales[[#This Row],[Qty]]`), as a file writes them.
 */
function formulaText(f: string, union = false, tables?: Tables, enclosing?: string): string {
  if (tables) {
    const r = validate(f, { union, tables: tables.list.map((t) => ({ name: t.name, columns: t.cols })), ...(enclosing ? { enclosing } : {}) })
    if (r.ok) return r.canonical
  }
  try {
    return canonical(parse(f, { union }))
  } catch {
    return f
  }
}

/** An error value as a values-only file stores it and a reader reads it back: a formula of that error. */
const ERROR_FORMULA: Partial<Record<string, string>> = { DIV0: '#DIV/0!', NA: '#N/A', NAME: '#NAME?', NUM: '#NUM!', REF: '#REF!', VALUE: '#VALUE!', ERROR: '#NULL!', SPILL: '#SPILL!' }

type Look = { style?: CellStyle; format?: NumberFormat }

function lookOf(book: BookInput, x: { style?: number; numberFormat?: NumberFormat; format?: NumberFormat }): Look {
  const out: Look = {}
  const style = normStyle(x.style === undefined ? undefined : book.styles?.[x.style])
  if (style) out.style = style
  // a file stores a code that every locale shows as a built-in format by that format's number
  const f = numberFormatOf(x)
  if (f !== undefined && f !== 'General' && f !== 0) out.format = typeof f === 'string' ? (BUILTIN_ID.get(f) ?? f) : f
  return out
}

/**
 * The entries of a book. With `values` (a values-only file and the book's result), a formula's cell and the
 * cells its array spilled into hold the value instead: a stored error is read back as a formula of it, a blank
 * result as 0, #CYCLE! as 0.
 */
export function entries(book: BookInput, values?: ComputeResult): Map<string, string> {
  const out = new Map<string, string>()
  const tables = tablesOf(book)
  const epoch = epochOf(book.dateSystem)
  book.sheets.forEach((s, i) => {
    out.set(`sheet ${i + 1}`, s.name)
    if (s.filtered) out.set(`filtered ${s.name}`, 'true')
    const res = values?.sheets[i].cells
    const cells = new Map<string, Record<string, unknown>>()
    // #CALC! has no formula spelling: the reader leaves it out
    const fromValue = (v: CellValue): Record<string, unknown> =>
      v === null ? { value: 0 } : typeof v === 'object' ? (v.error === 'CYCLE' ? { value: 0 } : ERROR_FORMULA[v.error] ? { formula: ERROR_FORMULA[v.error] } : {}) : { value: v }
    // a look a cell repeats from its row, else its column, is not its own
    const rowLines = lines(s.rows, 'r')
    const colLines = lines(s.columns, 'c')
    const inherited = (addr: string): Look => {
      const { row, col } = parseAddress(addr)
      const [r, k] = [entryAt(rowLines, row), entryAt(colLines, col)]
      return lookOf(book, { style: r?.style ?? k?.style, numberFormat: numberFormatOf(r) ?? numberFormatOf(k) })
    }
    for (const [addr, c] of Object.entries(s.cells) as [string, CellInput][]) {
      const own = lookOf(book, c)
      const x: Record<string, unknown> = stable(own) === stable(inherited(addr)) ? {} : { ...own }
      const { row, col } = parseAddress(addr)
      if (c.formula !== undefined) Object.assign(x, res ? fromValue(res[addr]?.value ?? null) : { formula: formulaText(c.formula, false, tables, tableAt(tables, i, row, col)?.name), array: c.array })
      else if (c.value !== undefined && c.value !== null) x.value = typeof c.value === 'object' ? dateToSerial(c.value.date) - epoch : c.value
      if (Object.keys(x).length) cells.set(addr, x)
    }
    if (res)
      for (const [addr, r] of Object.entries(res) as [string, CellResult][])
        if (r.spillOf !== undefined) cells.set(addr, { ...(cells.get(addr) ?? {}), ...fromValue(r.value) })
    for (const [addr, x] of cells) out.set(`cell ${s.name}!${addr}`, stable(x))
    // a line of size 0 is hidden in Excel
    for (const [lo, hi, e] of lines(s.columns, 'c')) for (let c = lo; c <= hi; c++) out.set(`column ${s.name}!${indexToCol(c)}`, stable({ width: e.width, hidden: e.hidden || e.width === 0 || undefined, ...lookOf(book, e) }))
    for (const [lo, hi, e] of lines(s.rows, 'r')) for (let r = lo; r <= hi; r++) out.set(`row ${s.name}!${r + 1}`, stable({ height: e.height, hidden: e.hidden || e.height === 0 || undefined, ...lookOf(book, e) }))
    if (s.merges?.length) out.set(`merges ${s.name}`, stable([...s.merges].sort()))
    // a table's columns: a calculated column's formula as stored text; a totals formula of its own (`custom`) and a label
    // the totals cell holds are left out (the cell holds them; Excel records a label for the text there when it saves)
    for (const t of s.tables ?? []) {
      const x = tables.byName.get(t.name.toLocaleUpperCase('ru-RU'))!
      const foot = (i: number) => (t.totals ? s.cells[`${indexToCol(x.c1 + i)}${x.r2 + 1}`]?.value : undefined)
      const own = (c: TableInput['columns'][number], i: number) => !c.totals || ('function' in c.totals ? c.totals.function === 'custom' : c.totals.label === foot(i))
      out.set(
        `table ${s.name}!${t.name}`,
        stable({
          ...t,
          columns: t.columns.map((c, i) => ({
            name: c.name,
            ...(c.formula !== undefined ? { formula: formulaText(c.formula, false, tables, t.name) } : {}),
            ...(own(c, i) ? {} : { totals: c.totals }),
          })),
        }),
      )
    }
  })
  for (const n of book.names ?? []) out.set(`name ${n.scope ?? ''}!${n.name.toUpperCase()}`, formulaText(n.ref, true, tables))
  out.set('locale', book.locale ?? 'en-US')
  out.set('date system', String(book.dateSystem ?? 1900))
  if (book.iteration) out.set('iteration', stable({ maxIterations: 100, maxChange: 0.001, ...book.iteration }))
  return out
}

/** Places where two books differ (`entries` of each). */
export function compareBooks(written: Map<string, string>, read: Map<string, string>): Difference[] {
  const out: Difference[] = []
  for (const place of new Set([...written.keys(), ...read.keys()])) {
    const [w, r] = [written.get(place), read.get(place)]
    if (w !== r) out.push({ place, ...(w !== undefined ? { written: w } : {}), ...(r !== undefined ? { read: r } : {}) })
  }
  return out
}

/** A value as both sides are compared: numbers to 15 significant digits, a blank result as 0. */
const plain = (v: CellValue | undefined): string => (v === undefined ? 'none' : v === null ? '0' : typeof v === 'number' ? String(Number(v.toPrecision(15))) : stable(v))

/** Cells where the values Excel stored differ from the engine's: formulas and the cells their arrays spilled into. */
export function compareCached(result: ComputeResult, cached: XlsxCached): Difference[] {
  const out: Difference[] = []
  result.sheets.forEach((s, i) => {
    const stored = cached.sheets[i]?.cells ?? {}
    const places = new Set([...Object.keys(stored), ...Object.entries(s.cells).flatMap(([a, c]) => (c.formula !== undefined || c.spillOf !== undefined ? [a] : []))])
    for (const addr of places) {
      const [engine, excel] = [plain(s.cells[addr]?.value), plain(stored[addr]?.value)]
      if (engine !== excel) out.push({ place: `${s.name}!${addr}`, written: engine, read: excel })
    }
  })
  return out
}
