/**
 * Excel's answers to structure operations (`test/excel/struct/<set>.json`, from `run.ts struct <set>`)
 * and the comparison with `editStructure`: what each cell shows (`=formula` or the value) in the read
 * rectangle of every sheet, the extra cells, the sheets and the names. Formulas are compared in the form
 * Excel writes references (`A1:A1048576` is `A:A`), with sheet names quoted.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import type { CellLook, ExcelStructAnswer, SheetLook, StoredStruct } from '../bench/coverage/struct.js'
import type { BookInput, CellInput, SheetInput } from '../src/contract.js'
import { isHiddenRow } from '../src/layout/hidden.js'
import { parseMerge } from '../src/layout/merges.js'
import { lineAt } from '../src/layout/spans.js'
import { formatAddress, indexToCol, parseAddress } from '../src/model.js'
import { cornersText, decode, encode, meets } from '../src/structure/area.js'
import { canonical, parse, sheetKey, walk } from '../src/syntax.js'
import { expandSheet } from './looks-book.js'
import { numberFormatOf, takenBook, type StoredBook } from '../bench/coverage/stored.js'
import type { StructCase } from '../bench/coverage/struct/cases.js'
import { editStructure } from '../src/structure/index.js'

const dir = new URL('./excel/struct/', import.meta.url)

export const STRUCT_SETS: (StoredStruct & { name: string })[] = existsSync(dir)
  ? readdirSync(dir)
      .filter((f) => f.endsWith('.json'))
      .sort()
      .map((f) => ({ name: f.slice(0, -5), ...(JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as StoredStruct) }))
  : []

/** Formula text (without `=`, a name's union too) with references as Excel writes them; as written when it does not parse. */
export function normalFormula(f: string): string {
  let ast
  try {
    ast = parse(f, { union: true })
  } catch {
    return f
  }
  walk(ast, (n) => {
    if (n.k === 'ref' || n.k === 'ref3d') n.ref = { ...n.ref, text: encode(decode(n.ref)) }
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
      const format = numberFormatOf(cell) ?? numberFormatOf(row) ?? numberFormatOf(col)
      const align = style?.alignment?.horizontal
      const l: CellLook = { ...(style?.font?.bold ? { bold: true } : {}), ...(style?.fill ? { fill: style.fill.color.toUpperCase() } : {}), ...(format !== undefined && format !== 'General' ? { format: String(format) } : {}), ...(align && align !== 'general' ? { align } : {}) }
      if (Object.keys(l).length) out.cells[addr] = l
    }
  // a hidden line shows no size, as Excel reads it; a row of height 0 is hidden
  const hidden = (k: string) => (out.hidden ??= []).includes(k) || out.hidden.push(k)
  for (let c = a.col; c <= b.col; c++) {
    const col = lineAt(sheet.columns, 'c', c)
    if (col?.hidden) hidden(indexToCol(c))
    else if (col?.width !== undefined) out.widths[indexToCol(c)] = shownWidth(col.width)
  }
  for (let r = a.row; r <= b.row; r++) {
    const row = lineAt(sheet.rows, 'r', r)
    if (row && isHiddenRow(row)) hidden(String(r + 1))
    else if (row?.height !== undefined) out.heights[String(r + 1)] = shownHeight(row.height)
  }
  const window = { r1: a.row, c1: a.col, r2: b.row, c2: b.col }
  for (const m of merges) if (meets(m, window) && !out.merges.includes(cornersText(m))) out.merges.push(cornersText(m))
}

/** A merged area of whole rows or columns as Excel writes its address (`20:21`, `J:K`), in cells (`A20:XFD21`). */
const fullLines = (m: string) => m.replace(/^(\d+):(\d+)$/, 'A$1:XFD$2').replace(/^([A-Z]+):([A-Z]+)$/, '$11:$21048576')

/** Differences of looks, as `place: Excel …, engine …` lines. */
function lookDifferences(name: string, excel: SheetLook, ours: SheetLook, widths: boolean): string[] {
  const out: string[] = []
  const diff = (what: string, x: Record<string, unknown>, y: Record<string, unknown>) => {
    for (const k of new Set([...Object.keys(x), ...Object.keys(y)])) if (JSON.stringify(x[k] ?? null) !== JSON.stringify(y[k] ?? null)) out.push(`'${name}' ${what} ${k}: Excel ${JSON.stringify(x[k] ?? null)}, engine ${JSON.stringify(y[k] ?? null)}`)
  }
  diff('look', excel.cells, ours.cells)
  if (widths) diff('width', excel.widths, ours.widths)
  diff('height', excel.heights, ours.heights)
  const [hx, hy] = [JSON.stringify([...(excel.hidden ?? [])].sort()), JSON.stringify([...(ours.hidden ?? [])].sort())]
  if (hx !== hy) out.push(`'${name}' hidden: Excel ${hx}, engine ${hy}`)
  const [x, y] = [excel.merges.map(fullLines).sort(), [...ours.merges].sort()]
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
    const found = book.sheets.find((x) => x.name === s.name)
    if (!found) continue
    // the cells of the sheet's looks as the entries they stand for
    const ours = expandSheet(found)
    // every cell of an old-kind array formula's area shows its formula, as Excel shows it
    const arrays = new Map<string, CellInput>()
    for (const c of Object.values(ours.cells)) {
      if (c.array === undefined) continue
      const [p, q = p] = c.array.split(':').map(parseAddress)
      for (let r = p.row; r <= q.row; r++) for (let k = p.col; k <= q.col; k++) arrays.set(formatAddress(r, k), c)
    }
    const addrs = new Set([...Object.keys(s.cells), ...Object.keys(ours.cells).filter(inRead), ...[...arrays.keys()].filter(inRead)])
    for (const addr of addrs) {
      const excel = s.cells[addr] === undefined ? '' : excelShown(s.cells[addr])
      const own = ours.cells[addr] ?? arrays.get(addr)
      const engine = own ? shown(own) : ''
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

/** The case's operation on its book, after the one taken before it. */
export function structureOf(book: StoredBook, c: StructCase) {
  const first = c.before ? editStructure(takenBook(book), c.before) : null
  if (first && !first.ok) throw new Error(`${c.id}: the engine refused the operation before: ${first.message}`)
  return editStructure(first ? first.book : takenBook(book), c.op)
}

export { formatAddress }
