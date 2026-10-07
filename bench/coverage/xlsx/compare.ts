/**
 * Comparing a book with the one read back from a file, and Excel's stored values with the engine's. Both sides
 * are flattened to entries by place (`cell S!A1`, `column S!B`, `name S!rate` …) holding plain values: a style by
 * its fields rather than its index, a formula as its canonical text, a date as its serial number, spans of rows
 * and columns line by line.
 */
import type { BookInput, CellInput, CellResult, CellStyle, CellValue, ComputeResult, NumberFormat } from '../../../src/contract.js'
import { entryAt, lines } from '../../../src/layout/spans.js'
import { dateToSerial, indexToCol, parseAddress } from '../../../src/model.js'
import { canonical, parse } from '../../../src/syntax.js'

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

function formulaText(f: string): string {
  try {
    return canonical(parse(f))
  } catch {
    return f
  }
}

/** An error value as a values-only file stores it and a reader reads it back: a formula of that error. */
const ERROR_FORMULA: Partial<Record<string, string>> = { DIV0: '#DIV/0!', NA: '#N/A', NAME: '#NAME?', NUM: '#NUM!', REF: '#REF!', VALUE: '#VALUE!', ERROR: '#NULL!', SPILL: '#SPILL!' }

type Look = { style?: CellStyle; format?: NumberFormat }

function lookOf(book: BookInput, x: { style?: number; format?: NumberFormat }): Look {
  const out: Look = {}
  const style = normStyle(x.style === undefined ? undefined : book.styles?.[x.style])
  if (style) out.style = style
  if (x.format !== undefined && x.format !== 'General') out.format = x.format
  return out
}

/**
 * The entries of a book. With `values` (a values-only file and the book's result), a formula's cell and the
 * cells its array spilled into hold the value instead: a stored error is read back as a formula of it, a blank
 * result as 0, #CYCLE! as 0.
 */
export function entries(book: BookInput, values?: ComputeResult): Map<string, string> {
  const out = new Map<string, string>()
  book.sheets.forEach((s, i) => {
    out.set(`sheet ${i + 1}`, s.name)
    const res = values?.sheets[i].cells
    const cells = new Map<string, Record<string, unknown>>()
    // #CALC! has no formula spelling: the reader leaves it out
    const fromValue = (v: CellValue): Record<string, unknown> =>
      v === null ? { value: 0 } : typeof v === 'object' ? (v.error === 'CYCLE' ? { value: 0 } : ERROR_FORMULA[v.error] ? { formula: ERROR_FORMULA[v.error] } : {}) : { value: v }
    // a cell shows its own look, else its row's, else its column's: a look a cell repeats from them is not its own
    const rowLines = lines(s.rows, 'r')
    const colLines = lines(s.columns, 'c')
    const inherited = (addr: string): Look => {
      const { row, col } = parseAddress(addr)
      const [r, c] = [entryAt(rowLines, row), entryAt(colLines, col)]
      return lookOf(book, { style: r?.style ?? c?.style, format: r?.format ?? c?.format })
    }
    for (const [addr, c] of Object.entries(s.cells) as [string, CellInput][]) {
      const own = lookOf(book, c)
      const x: Record<string, unknown> = stable(own) === stable(inherited(addr)) ? {} : { ...own }
      if (c.formula !== undefined) Object.assign(x, res ? fromValue(res[addr]?.value ?? null) : { formula: formulaText(c.formula) })
      else if (c.value !== undefined && c.value !== null) x.value = typeof c.value === 'object' ? dateToSerial(c.value.date) : c.value
      if (Object.keys(x).length) cells.set(addr, x)
    }
    if (res)
      for (const [addr, r] of Object.entries(res) as [string, CellResult][])
        if (r.spillOf !== undefined) cells.set(addr, { ...(cells.get(addr) ?? {}), ...fromValue(r.value) })
    for (const [addr, x] of cells) out.set(`cell ${s.name}!${addr}`, stable(x))
    for (const [lo, hi, e] of lines(s.columns, 'c')) for (let c = lo; c <= hi; c++) out.set(`column ${s.name}!${indexToCol(c)}`, stable({ width: e.width, ...lookOf(book, e) }))
    for (const [lo, hi, e] of lines(s.rows, 'r')) for (let r = lo; r <= hi; r++) out.set(`row ${s.name}!${r + 1}`, stable({ height: e.height, ...lookOf(book, e) }))
    if (s.merges?.length) out.set(`merges ${s.name}`, stable([...s.merges].sort()))
  })
  for (const n of book.names ?? []) out.set(`name ${n.scope ?? ''}!${n.name.toUpperCase()}`, formulaText(n.ref))
  out.set('locale', book.locale ?? 'en-US')
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
export function compareCached(result: ComputeResult, cached: { name: string; cells: Record<string, CellValue> }[]): Difference[] {
  const out: Difference[] = []
  result.sheets.forEach((s, i) => {
    const stored = cached[i]?.cells ?? {}
    const places = new Set([...Object.keys(stored), ...Object.entries(s.cells).flatMap(([a, c]) => (c.formula !== undefined || c.spillOf !== undefined ? [a] : []))])
    for (const addr of places) {
      const [engine, excel] = [plain(s.cells[addr]?.value), plain(stored[addr])]
      if (engine !== excel) out.push({ place: `${s.name}!${addr}`, written: engine, read: excel })
    }
  })
  return out
}
