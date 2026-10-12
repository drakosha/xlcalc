/**
 * A small case book as xlsx and the reading of its rectangles, shared by the steps `spill` and `workbook`: the
 * book's formulas as dynamic array formulas, its names, merged areas and looks, what the file holds beyond the
 * model (`FileParts`: hidden rows and columns, filters, the 1904 date system, iterative calculation, cached
 * values), and the mirror sheets `WT`, `WN`, `WX` whose formulas never give an error.
 *
 * Reading: the `value` of a range with errors in it is unsafe (errors drop out of the list, a two-dimensional
 * read crashed Excel), so every cell of a rectangle has three mirror formulas, a type code (-5 blank, 0
 * number, -1 TRUE, -2 FALSE, -3 text, > 0 `ERROR.TYPE`), the number and the text, and those are read.
 */
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { readFileSync, writeFileSync } from 'node:fs'
import { anchorArray, markDynamic, READ_COLUMNS } from './book.js'
import { runWatched, settleRejected, typedValue } from './excel.js'
import type { Val } from './formulajs.js'
import { xl } from './sets/arrays/common.js'
import { xlpm } from './sets/let/xlpm.js'
import type { BookInput, NumberFormat, RangeRef } from '../../src/contract.js'
import { lines, parseSpan } from '../../src/layout/spans.js'
import { formatAddress } from '../../src/model.js'
import { parseRange, rectCells } from '../../src/structure/area.js'
import { quoteSheet } from '../../src/syntax.js'
import { numberFormatOf } from './stored.js'

export const MIRRORS = ['WT', 'WN', 'WX'] as const

/** What Excel shows in a case's rectangles. */
export interface RectsAnswer {
  /** Non-blank cells of the case's rectangles (`S!A1`). */
  values: Record<string, Val>
  /** Formulas as Excel shows them, for the cells that have one (`=SEQUENCE(3)`). */
  formulas: Record<string, string>
}

/** A sheet's parts of a file that the book does not write: hidden lines, an autofilter with its criteria. */
export interface SheetParts {
  /** Rows (`5`, `5:9`) and columns (`C`, `B:D`) hidden. */
  hiddenRows?: string[]
  hiddenColumns?: string[]
  /** Rows (`5`, `5:9`) of height 0, which ExcelJS writes without their height. */
  zeroRows?: string[]
  /**
   * An autofilter over `range`: the values each column (0 from its left) keeps; the rows it hides are among
   * `hiddenRows`. The sheet is marked as filtered (`<sheetPr filterMode="1">`) unless `mode` is false.
   */
  filter?: { range: string; keep: Record<number, string[]>; mode?: false }
}

/** What the file holds beyond the book's model. */
export interface FileParts {
  sheets?: Record<string, SheetParts>
  /** `<workbookPr date1904="1">`. */
  date1904?: true
  /** `<calcPr iterate="1" iterateCount iterateDelta>`; Excel's defaults (100, 0.001) where left out. */
  iterate?: { count?: number; delta?: number }
  /** Values the file stores for formulas (`S!A1`), as Excel stores what it computed last. */
  cached?: Record<string, number>
  /**
   * Formulas written as ordinary formulas, as Excel 365 saves a formula that does not spill, but for the cells listed
   * (`S!A1`), written as dynamic array formulas; without it every formula is a dynamic array formula.
   */
  plain?: { arrays?: string[] }
  /**
   * The mirror sheets' names, `WT`, `WN`, `WX` without it. Their formulas depend on the cells read. With them sorting
   * before a cycle's sheet, Excel's first computation on opening did one round of iteration less (the cause was not
   * found: one sheet of a dependent before the cycle's did not do it), so the set `iterate` names them to sort last.
   */
  mirrors?: readonly [string, string, string]
}

/**
 * A formula in the engine's spelling as Excel stores it: newer functions prefixed, LET's names as `_xlpm.x`, `E2#` as
 * `ANCHORARRAY(E2)`, the operator `@x` as `SINGLE(x)`.
 */
export const fileFormula = (f: string) => anchorArray(single(xl(xlpm(f))))

/** The end of the run quoted by `q` that starts at `i` (`""` and `''` inside it are kept). */
function quoted(f: string, i: number, q: string): number {
  let j = i + 1
  while (j < f.length && !(f[j] === q && f[j + 1] !== q)) j += f[j] === q ? 2 : 1
  return j + 1
}

/** The end of the parenthesis that opens at `i`. */
function closing(f: string, i: number): number {
  let depth = 0
  for (let j = i; j < f.length; j++) {
    if (f[j] === '"' || f[j] === "'") j = quoted(f, j, f[j]) - 1
    else if (f[j] === '(') depth++
    else if (f[j] === ')' && --depth === 0) return j + 1
  }
  return f.length
}

/** `@x` as `_xlfn.SINGLE(x)`: x a parenthesis, a call, a string, an array constant, or a reference, a name or a literal. */
function single(f: string): string {
  let out = ''
  for (let i = 0; i < f.length; ) {
    const ch = f[i]
    if (ch === '"') {
      const j = quoted(f, i, ch)
      out += f.slice(i, j)
      i = j
      continue
    }
    if (ch !== '@') {
      out += ch
      i++
      continue
    }
    let j = i + 1
    if (f[j] === '"') j = quoted(f, j, '"')
    else if (f[j] === '{') j = f.indexOf('}', j) + 1
    else {
      if (f[j] === "'") j = quoted(f, j, "'")
      while (j < f.length && /[\w.$:!#]/.test(f[j])) j++
      if (f[j] === '(') j = closing(f, j)
    }
    out += `_xlfn.SINGLE(${single(f.slice(i + 1, j))})`
    i = j
  }
  return out
}

/** A cell's format as ExcelJS writes it: a code; a built-in format by its number is written by the display step only. */
export function formatCode(format: NumberFormat): string {
  if (typeof format === 'number') throw new Error(`the built-in format ${format} by its number: only the display step writes it`)
  return format
}

/** A rectangle's cells row by row. */
export function cellsOf(range: string): { rows: number; cols: number; addrs: string[] } {
  const x = parseRange(range)!
  return { rows: x.r2 - x.r1 + 1, cols: x.c2 - x.c1 + 1, addrs: [...rectCells(x)] }
}

/** Where each rectangle's mirrors are: stacked from row 1 of the mirror sheets. */
export function mirrorRanges(read: RangeRef[]): string[] {
  let row = 0
  return read.map((g) => {
    const { rows, cols } = cellsOf(g.range)
    const out = `A${row + 1}:${formatAddress(row + rows - 1, cols - 1)}`
    row += rows
    return out
  })
}

/** The read columns, a blank cell typed -5 apart. */
const MIRROR: Record<(typeof MIRRORS)[number], (x: string) => string> = {
  WT: (x) => `IF(ISBLANK(${x}),-5,${READ_COLUMNS.type(x)})`,
  WN: READ_COLUMNS.num,
  WX: READ_COLUMNS.text,
}

/** The mirror formulas over a rectangle's top-left cell, its sheet's name quoted: the tables step writes them after its steps, over the cells where they are then. */
export function mirrorFormulas(g: RangeRef): string[] {
  const at = `'${g.sheet.replaceAll("'", "''")}'!${g.range.split(':')[0]}`
  return MIRRORS.map((m) => `=${MIRROR[m](at)}`)
}

/** A formula cell as ExcelJS writes it: an array formula over `ref` (a dynamic one where `markDynamic` marks it), else a plain one. */
export const formulaValue = (formula: string, ref?: string): ExcelJS.CellFormulaValue =>
  ({ formula: fileFormula(formula), ...(ref ? { shareType: 'array', ref } : {}) }) as ExcelJS.CellFormulaValue

/** The book's names with their scope, in place of ExcelJS's model, a cell matrix without `#`, formulas or scope. */
export function writeNames(wb: ExcelJS.Workbook, book: BookInput): void {
  const index = new Map(book.sheets.map((s, i) => [s.name, i]))
  const names = (book.names ?? []).map((n) => ({ name: n.name, ranges: [fileFormula(n.ref)], ...(n.scope !== undefined ? { localSheetId: index.get(n.scope) } : {}) }))
  Object.defineProperty(wb.definedNames, 'model', { get: () => names })
}

/** The case's book as xlsx: its formulas as dynamic array formulas, its names, merged areas, looks and file parts, then the mirror sheets. */
export async function writeCaseBook(book: BookInput, read: RangeRef[], file: string, parts: FileParts = {}): Promise<void> {
  const wb = new ExcelJS.Workbook()
  if (parts.date1904) wb.properties.date1904 = true
  for (const s of book.sheets) {
    const ws = wb.addWorksheet(s.name)
    for (const [addr, cell] of Object.entries(s.cells)) {
      const x = ws.getCell(addr)
      const result = parts.cached?.[`${s.name}!${addr}`]
      // an old-kind array formula over its area; a dynamic array formula, or a plain one where `plain` says
      const array = cell.array !== undefined || !parts.plain || parts.plain.arrays?.includes(`${s.name}!${addr}`)
      if (cell.formula !== undefined) x.value = { ...formulaValue(cell.formula, array ? (cell.array ?? addr) : undefined), ...(result === undefined ? {} : { result }) } as ExcelJS.CellValue
      else if (cell.value !== null && typeof cell.value === 'object') x.value = new Date(Date.parse(`${cell.value.date}${cell.value.date.includes('T') ? '' : 'T00:00:00'}Z`))
      else if (cell.value !== undefined) x.value = cell.value
      const numberFormat = numberFormatOf(cell)
      if (numberFormat !== undefined) x.numFmt = formatCode(numberFormat)
      const fill = cell.style === undefined ? undefined : book.styles![cell.style].fill
      if (fill) x.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${fill.color.slice(1)}` } }
    }
    for (const [lo, hi, e] of lines(s.columns, 'c')) for (let c = lo; c <= hi; c++) if (e.width !== undefined) ws.getColumn(c + 1).width = e.width
    for (const m of s.merges ?? []) ws.mergeCells(m)
    for (const k of parts.sheets?.[s.name]?.hiddenColumns ?? []) {
      const [lo, hi] = parseSpan(k, 'c')!
      for (let c = lo; c <= hi; c++) ws.getColumn(c + 1).hidden = true
    }
    const filter = parts.sheets?.[s.name]?.filter
    if (filter) ws.autoFilter = filter.range
  }
  const at = mirrorRanges(read)
  for (const [i, m] of MIRRORS.entries()) {
    const ws = wb.addWorksheet(parts.mirrors?.[i] ?? m)
    read.forEach((g, i) => {
      const src = cellsOf(g.range)
      const dst = cellsOf(at[i])
      src.addrs.forEach((a, k) => (ws.getCell(dst.addrs[k]).value = { formula: MIRROR[m](`${quoteSheet(g.sheet)}!${a}`) } as ExcelJS.CellValue))
    })
  }
  writeNames(wb, book)
  await wb.xlsx.writeFile(file)
  await markDynamic(file, book.sheets.length, (i, addr) => book.sheets[i]?.cells[addr]?.array !== undefined)
  if (parts.sheets || parts.iterate) await writeParts(file, book, parts)
}

/** The file parts ExcelJS does not write: hidden rows without cells, rows of height 0, an autofilter's criteria, iterative calculation. */
export async function writeParts(file: string, book: BookInput, parts: FileParts) {
  const zip = await JSZip.loadAsync(readFileSync(file))
  // the rows (1-based) of spans `5`, `5:9`
  const rowsOf = (spans: string[] = []) =>
    new Set(spans.flatMap((k) => {
      const [lo, hi] = parseSpan(k, 'r')!
      return Array.from({ length: hi - lo + 1 }, (_, j) => lo + j + 1)
    }))
  for (const [i, s] of book.sheets.entries()) {
    const p = parts.sheets?.[s.name]
    if (!p?.hiddenRows && !p?.filter && !p?.zeroRows) continue
    const path = `xl/worksheets/sheet${i + 1}.xml`
    let xml = await zip.file(path)!.async('string')
    xml = markRows(markRows(xml, rowsOf(p.hiddenRows), 'hidden="1"'), rowsOf(p.zeroRows), 'ht="0" customHeight="1"')
    if (p.filter) {
      const cols = Object.entries(p.filter.keep).map(([col, vals]) => `<filterColumn colId="${col}"><filters>${vals.map((v) => `<filter val="${v}"/>`).join('')}</filters></filterColumn>`)
      xml = xml.replace(/<autoFilter ref="([^"]+)"\/>/, `<autoFilter ref="$1">${cols.join('')}</autoFilter>`)
      if (p.filter.mode !== false) xml = xml.replace(/(<worksheet[^>]*>)/, '$1<sheetPr filterMode="1"/>')
    }
    zip.file(path, xml)
  }
  if (parts.iterate) {
    const xml = await zip.file('xl/workbook.xml')!.async('string')
    const { count = 100, delta = 0.001 } = parts.iterate
    zip.file('xl/workbook.xml', xml.replace(/<calcPr ([^/]*)\/>/, `<calcPr $1 iterate="1" iterateCount="${count}" iterateDelta="${delta}"/>`))
  }
  writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
}

/** `attrs` on the rows given (1-based); a row ExcelJS left out (no cells) is written empty. */
function markRows(xml: string, rows: Set<number>, attrs: string): string {
  if (!rows.size) return xml
  const todo = new Set(rows)
  let out = xml.replace(/<row r="(\d+)"/g, (m, r: string) => (todo.delete(Number(r)) ? `${m} ${attrs}` : m))
  const empty = (r: number) => `<row r="${r}" ${attrs}/>`
  for (const r of [...todo].sort((a, b) => a - b)) {
    // before the first row past it, else at the end of the sheet's data
    const next = [...out.matchAll(/<row r="(\d+)"/g)].find((m) => Number(m[1]) > r)
    out = next ? out.slice(0, next.index) + empty(r) + out.slice(next.index) : out.replace(/<sheetData\/>|<\/sheetData>/, (m) => (m === '<sheetData/>' ? `<sheetData>${empty(r)}</sheetData>` : `${empty(r)}</sheetData>`))
  }
  return out
}

/** Excel did not load a case's book: it showed "We found a problem with some content", which was declined. */
export class BookRejected extends Error {}

/**
 * A case's book in Excel's sandbox, the AppleScript `script` over it (arguments: the path, the rectangles as lines
 * `sheet<TAB>range<TAB>mirror range`, then `more`): its output. The repair window is declined and `BookRejected`
 * thrown. Another window that is not a book stops the run and is left open (nothing is clicked); so does Excel not
 * answering in 120 s.
 */
export async function runOnBook(id: string, script: string, book: BookInput, read: RangeRef[], more: string[], parts?: FileParts): Promise<string> {
  const at = mirrorRanges(read)
  return runWatched<string>({
    script,
    label: id,
    folder: 'fcase',
    book: (file) => writeCaseBook(book, read, file, parts),
    args: () => [read.map((g, i) => `${g.sheet}\t${g.range}\t${at[i]}`).join('\n'), ...more],
    timeout: 120_000,
    poll: 1000,
    after: 'case',
    fatal: true,
    repair: async (stop) => {
      await stop()
      await settleRejected()
      throw new BookRejected(`${id}: Excel did not load the book (repair window declined)`)
    },
    done: (out) => out,
  })
}

/**
 * A value typed into a cell as the scripts take it, a kind and its text: x clears the cell, n a number, s a text, f a
 * formula as Excel 365 enters it, l a formula in the user's language.
 */
export function editOf(v: null | number | string | { formula: string } | { local: string }): [kind: 'x' | 'n' | 's' | 'f' | 'l', value: string] {
  if (v === null) return ['x', '']
  if (typeof v === 'number') return ['n', String(v)]
  if (typeof v === 'string') return ['s', v]
  return 'local' in v ? ['l', v.local] : ['f', `=${v.formula}`]
}

/** The steps of a case Excel refused, by their index: their errors (an empty one or none where Excel took the step); undefined when it took all. */
export function refusedOf(errors: (string | null | undefined)[]): Record<number, string> | undefined {
  let out: Record<number, string> | undefined
  errors.forEach((e, i) => e && ((out ??= {})[i] = e))
  return out
}

/** Each cell of the rectangles read, row by row: its key (`S!A1`), its rectangle's index, its row and column there. */
export function eachRectCell(read: RangeRef[], f: (key: string, i: number, r: number, j: number) => void): void {
  read.forEach((g, i) => {
    const { cols, addrs } = cellsOf(g.range)
    addrs.forEach((a, k) => f(`${g.sheet}!${a}`, i, Math.floor(k / cols), k % cols))
  })
}

/** A range's `value` or `formula` as AppleScript returns it (a cell alone is not a list) as rows. */
export const grid = (x: unknown): unknown[][] => (Array.isArray(x) ? (Array.isArray(x[0]) ? (x as unknown[][]) : [x]) : [[x]])

/**
 * The answer from the read of each rectangle: its mirrors' types, numbers and texts, and its formulas; with `withF2`
 * the formulas as Excel 365 shows them too (`formulas2`), read fifth.
 */
export function answerOf(read: RangeRef[], raw: [unknown, unknown, unknown, unknown, unknown?][], withF2 = false): RectsAnswer & { formulas2?: Record<string, string> } {
  const out: RectsAnswer & { formulas2?: Record<string, string> } = { values: {}, formulas: {}, ...(withF2 ? { formulas2: {} } : {}) }
  const grids = raw.map((x) => x.map(grid))
  eachRectCell(read, (key, i, r, j) => {
    const [t, nums, strs, fs, f2s] = grids[i]
    const type = Number(t[r][j])
    const f = String(fs[r][j])
    if (f.startsWith('=')) out.formulas[key] = f
    if (out.formulas2 && f.startsWith('=')) out.formulas2[key] = String(f2s[r][j])
    if (type === -5) return
    out.values[key] = typedValue(type, nums[r][j], strs[r][j])
  })
  return out
}
