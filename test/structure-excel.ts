/**
 * Excel's answers to structure operations (`test/excel/struct/<set>.json`, from `run.ts struct <set>`)
 * and the comparison with `restructure`: what each cell shows (`=formula` or the value) in the read
 * rectangle of every sheet, the extra cells, the sheets and the names. Formulas are compared in the form
 * Excel writes references (`A1:A1048576` is `A:A`), with sheet names quoted.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import type { CellLook, ExcelStructAnswer, SheetLook, StoredStruct } from '../bench/coverage/struct.js'
import type { BookInput, CellInput, SheetInput } from '../src/contract.js'
import { mergeText, parseMerge } from '../src/layout/merges.js'
import { lineAt } from '../src/layout/spans.js'
import { formatAddress, indexToCol, parseAddress } from '../src/model.js'
import { meets } from '../src/structure/area.js'
import { decode, encode } from '../src/structure/area.js'
import { canonical, parse, sheetKey, walk } from '../src/syntax.js'

const dir = new URL('./excel/struct/', import.meta.url)

export const STRUCT_SETS: (StoredStruct & { name: string })[] = existsSync(dir)
  ? readdirSync(dir)
      .filter((f) => f.endsWith('.json'))
      .sort()
      .map((f) => ({ name: f.slice(0, -5), ...(JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as StoredStruct) }))
  : []

/** Formula text (without `=`) with references as Excel writes them; as written when it does not parse. */
export function normalFormula(f: string): string {
  let ast
  try {
    ast = parse(f)
  } catch {
    return f
  }
  walk(ast, (n) => {
    if (n.k === 'ref') n.ref = { ...n.ref, text: encode(decode(n.ref)) }
  })
  return canonical(ast)
}

/** What Excel's `formula` of a cell shows for our cell. */
function shown(c: CellInput): string {
  if (c.formula !== undefined) return `=${normalFormula(c.formula)}`
  const v = c.value
  if (v === undefined || v === null) return ''
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  return typeof v === 'object' ? v.date : String(v)
}

const excelShown = (text: string) => (text.startsWith('=') ? `=${normalFormula(text.slice(1))}` : text)

/** Excel's `Sheet!name` as scope and name. */
function nameKey(excel: string): string {
  const i = excel.lastIndexOf('!')
  if (i < 0) return `|${excel.toUpperCase()}`
  const scope = excel.slice(0, i).replace(/^'(.*)'$/, '$1').replaceAll("''", "'")
  return `${sheetKey(scope)}|${excel.slice(i + 1).toUpperCase()}`
}

/**
 * Sizes as Excel 16 for Mac reports them with Calibri 11: a column width of w characters (xlsx) as (6w − 5)/6,
 * the width without the padding at 6 pixels a digit; a row height in points as pixels, 4/3 of it, to a quarter
 * (30 is 40, 24 is 32, 20 is 26.75).
 */
const shownWidth = (w: number) => Math.round(((6 * w - 5) / 6) * 100) / 100
const shownHeight = (h: number) => Math.round((h * 16) / 3) / 4

/**
 * The look of a sheet of the book in the rectangles of `look` (`A1:J24,…`), as Excel reads it: a cell's style and
 * format, else its row's, else its column's; a cell of a merged area shows the area's top-left cell, as the area is
 * drawn (deleting row 1,048,576 keeps an area that reaches into it, `edge-layout`).
 */
export function lookOf(book: BookInput, sheet: SheetInput, look: string): SheetLook {
  const out: SheetLook = { cells: {}, widths: {}, heights: {}, merges: [] }
  for (const rect of look.split(',')) lookIn(book, sheet, rect, out)
  return out
}

function lookIn(book: BookInput, sheet: SheetInput, rect: string, out: SheetLook): void {
  const [a, b] = rect.split(':').map(parseAddress)
  const merges = (sheet.merges ?? []).map((m) => parseMerge(m)!)
  for (let r = a.row; r <= b.row; r++)
    for (let c = a.col; c <= b.col; c++) {
      const addr = formatAddress(r, c)
      const m = merges.find((x) => r >= x.r1 && r <= x.r2 && c >= x.c1 && c <= x.c2)
      const [tr, tc] = m ? [m.r1, m.c1] : [r, c]
      const [cell, row, col] = [sheet.cells[formatAddress(tr, tc)], lineAt(sheet.rows, 'r', tr), lineAt(sheet.columns, 'c', tc)]
      const style = book.styles?.[cell?.style ?? row?.style ?? col?.style ?? -1]
      const format = cell?.format ?? row?.format ?? col?.format
      const l: CellLook = { ...(style?.font?.bold ? { bold: true } : {}), ...(style?.fill ? { fill: style.fill.color.toUpperCase() } : {}), ...(format !== undefined && format !== 'General' ? { format } : {}) }
      if (Object.keys(l).length) out.cells[addr] = l
    }
  for (let c = a.col; c <= b.col; c++) {
    const w = lineAt(sheet.columns, 'c', c)?.width
    if (w !== undefined) out.widths[indexToCol(c)] = shownWidth(w)
  }
  for (let r = a.row; r <= b.row; r++) {
    const h = lineAt(sheet.rows, 'r', r)?.height
    if (h !== undefined) out.heights[String(r + 1)] = shownHeight(h)
  }
  const window = { r1: a.row, c1: a.col, r2: b.row, c2: b.col }
  for (const m of merges) if (meets(m, window) && !out.merges.includes(mergeText(m))) out.merges.push(mergeText(m))
}

/** Differences of looks, as `place: Excel …, engine …` lines. */
function lookDifferences(name: string, excel: SheetLook, ours: SheetLook, widths: boolean): string[] {
  const out: string[] = []
  const diff = (what: string, x: Record<string, unknown>, y: Record<string, unknown>) => {
    for (const k of new Set([...Object.keys(x), ...Object.keys(y)])) if (JSON.stringify(x[k] ?? null) !== JSON.stringify(y[k] ?? null)) out.push(`'${name}' ${what} ${k}: Excel ${JSON.stringify(x[k] ?? null)}, engine ${JSON.stringify(y[k] ?? null)}`)
  }
  diff('look', excel.cells, ours.cells)
  if (widths) diff('width', excel.widths, ours.widths)
  diff('height', excel.heights, ours.heights)
  const [x, y] = [[...excel.merges].sort(), [...ours.merges].sort()]
  if (JSON.stringify(x) !== JSON.stringify(y)) out.push(`'${name}' merges: Excel ${JSON.stringify(x)}, engine ${JSON.stringify(y)}`)
  return out
}

/** Differences between a book and Excel's answer, as `place: Excel … , engine …` lines; `look` for a styled book, column widths unless `widths` is false. */
export function differences(book: BookInput, a: ExcelStructAnswer, read: string, look?: string, widths = true): string[] {
  const out: string[] = []
  const [r0, r1] = read.split(':').map(parseAddress)
  const inRead = (addr: string) => {
    const p = parseAddress(addr)
    return p.row >= r0.row && p.row <= r1.row && p.col >= r0.col && p.col <= r1.col
  }
  const names = book.sheets.map((s) => s.name)
  if (JSON.stringify(names) !== JSON.stringify(a.sheets.map((s) => s.name))) out.push(`sheets: Excel ${JSON.stringify(a.sheets.map((s) => s.name))}, engine ${JSON.stringify(names)}`)
  for (const s of a.sheets) {
    const ours = book.sheets.find((x) => x.name === s.name)
    if (!ours) continue
    const addrs = new Set([...Object.keys(s.cells), ...Object.keys(ours.cells).filter(inRead)])
    for (const addr of addrs) {
      const excel = s.cells[addr] === undefined ? '' : excelShown(s.cells[addr])
      const engine = ours.cells[addr] ? shown(ours.cells[addr]) : ''
      if (excel !== engine) out.push(`'${s.name}'!${addr}: Excel ${JSON.stringify(excel)}, engine ${JSON.stringify(engine)}`)
    }
    // Excel's used range must lie in the read rectangle unless the case reads extra cells
    const used = s.used.replaceAll('$', '').split(':')
    if (!a.extra.length && !used.every((x) => /^[A-Z]+\d+$/.test(x) && inRead(x))) out.push(`'${s.name}': Excel's used range ${s.used} is outside ${read}`)
    if (s.look && look) out.push(...lookDifferences(s.name, s.look, lookOf(book, ours, look), widths))
  }
  for (const e of a.extra) {
    const c = book.sheets.find((x) => x.name === e.sheet)?.cells[e.cell]
    const engine = c ? shown(c) : ''
    if (excelShown(e.text) !== engine) out.push(`'${e.sheet}'!${e.cell}: Excel ${JSON.stringify(e.text)}, engine ${JSON.stringify(engine)}`)
  }
  const excelNames = new Map(a.names.filter((n) => !n.name.startsWith('_xlfn.')).map((n) => [nameKey(n.name), normalFormula(n.ref.slice(1))]))
  const ourNames = new Map((book.names ?? []).map((n) => [`${n.scope === undefined ? '' : sheetKey(n.scope)}|${n.name.toUpperCase()}`, normalFormula(n.ref)]))
  for (const k of new Set([...excelNames.keys(), ...ourNames.keys()]))
    if (excelNames.get(k) !== ourNames.get(k)) out.push(`name ${k}: Excel ${JSON.stringify(excelNames.get(k) ?? null)}, engine ${JSON.stringify(ourNames.get(k) ?? null)}`)
  return out
}

/** Number of formulas Excel showed in a case (for the report). */
export const formulaCount = (a: ExcelStructAnswer) => a.sheets.reduce((n, s) => n + Object.values(s.cells).filter((t) => t.startsWith('=')).length, 0) + a.names.filter((n) => !n.name.startsWith('_xlfn.')).length

export { formatAddress }
