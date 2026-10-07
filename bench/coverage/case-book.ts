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
import { spawn } from 'node:child_process'
import { mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { markDynamic } from './book.js'
import { dialogText, excelDialogs } from './excel.js'
import type { Val } from './formulajs.js'
import { xl } from './sets/arrays/common.js'
import { xlpm } from './sets/let/xlpm.js'
import type { BookInput, NumberFormat, RangeRef } from '../../src/contract.js'
import { lines, parseSpan } from '../../src/layout/spans.js'
import { formatAddress, parseAddress } from '../../src/model.js'
import { quoteSheet } from '../../src/syntax.js'

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

/** A formula in the engine's spelling as Excel stores it: newer functions prefixed, LET's names as `_xlpm.x`, `E2#` as `ANCHORARRAY(E2)`. */
export const fileFormula = (f: string) => xl(xlpm(f)).replace(/((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?\$?[A-Z]{1,3}\$?\d+#/g, (m) => `_xlfn.ANCHORARRAY(${m.slice(0, -1)})`)

const ERROR_TYPE: Record<number, Extract<Val, { t: 'e' }>['v']> = { 1: 'NULL', 2: 'DIV0', 3: 'VALUE', 4: 'REF', 5: 'NAME', 6: 'NUM', 7: 'NA', 8: 'GETTING_DATA', 9: 'SPILL', 14: 'CALC' }

/** A cell's format as ExcelJS writes it: a code; a built-in format by its number is written by the display step only. */
export function formatCode(format: NumberFormat): string {
  if (typeof format === 'number') throw new Error(`the built-in format ${format} by its number: only the display step writes it`)
  return format
}

/** A rectangle's cells row by row. */
export function cellsOf(range: string): { rows: number; cols: number; addrs: string[] } {
  const [a, b = a] = range.split(':').map(parseAddress)
  const addrs: string[] = []
  for (let r = a.row; r <= b.row; r++) for (let c = a.col; c <= b.col; c++) addrs.push(formatAddress(r, c))
  return { rows: b.row - a.row + 1, cols: b.col - a.col + 1, addrs }
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

const MIRROR: Record<(typeof MIRRORS)[number], (x: string) => string> = {
  WT: (x) => `IF(ISBLANK(${x}),-5,IF(ISERROR(${x}),ERROR.TYPE(${x}),IF(ISNUMBER(${x}),0,IF(ISLOGICAL(${x}),IF(${x},-1,-2),IF(ISTEXT(${x}),-3,-4)))))`,
  WN: (x) => `IF(ISNUMBER(${x}),${x},0)`,
  WX: (x) => `IF(ISTEXT(${x}),${x},"")`,
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
      const array = !parts.plain || parts.plain.arrays?.includes(`${s.name}!${addr}`)
      if (cell.formula !== undefined) x.value = { formula: fileFormula(cell.formula), ...(array ? { shareType: 'array', ref: addr } : {}), ...(result === undefined ? {} : { result }) } as ExcelJS.CellValue
      else if (cell.value !== null && typeof cell.value === 'object') x.value = new Date(Date.parse(`${cell.value.date}${cell.value.date.includes('T') ? '' : 'T00:00:00'}Z`))
      else if (cell.value !== undefined) x.value = cell.value
      if (cell.format !== undefined) x.numFmt = formatCode(cell.format)
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
  // ExcelJS keeps names in a cell matrix (no `#`): its model is replaced by the case's names
  const names = (book.names ?? []).map((n) => ({ name: n.name, ranges: [fileFormula(n.ref)] }))
  Object.defineProperty(wb.definedNames, 'model', { get: () => names })
  await wb.xlsx.writeFile(file)
  await markDynamic(file, book.sheets.length)
  if (parts.sheets || parts.iterate) await writeParts(file, book, parts)
}

/** The file parts ExcelJS does not write: hidden rows without cells, an autofilter's criteria, iterative calculation. */
export async function writeParts(file: string, book: BookInput, parts: FileParts) {
  const zip = await JSZip.loadAsync(readFileSync(file))
  for (const [i, s] of book.sheets.entries()) {
    const p = parts.sheets?.[s.name]
    if (!p?.hiddenRows && !p?.filter) continue
    const path = `xl/worksheets/sheet${i + 1}.xml`
    let xml = await zip.file(path)!.async('string')
    const hidden = new Set((p.hiddenRows ?? []).flatMap((k) => {
      const [lo, hi] = parseSpan(k, 'r')!
      return Array.from({ length: hi - lo + 1 }, (_, j) => lo + j + 1)
    }))
    xml = hideRows(xml, hidden)
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

/** `hidden="1"` on the rows given (1-based); a row ExcelJS left out (no cells) is written empty. */
function hideRows(xml: string, rows: Set<number>): string {
  if (!rows.size) return xml
  const todo = new Set(rows)
  let out = xml.replace(/<row r="(\d+)"/g, (m, r: string) => (todo.delete(Number(r)) ? `${m} hidden="1"` : m))
  const empty = (r: number) => `<row r="${r}" hidden="1"/>`
  for (const r of [...todo].sort((a, b) => a - b)) {
    // before the first row past it, else at the end of the sheet's data
    const next = [...out.matchAll(/<row r="(\d+)"/g)].find((m) => Number(m[1]) > r)
    out = next ? out.slice(0, next.index) + empty(r) + out.slice(next.index) : out.replace(/<sheetData\/>|<\/sheetData>/, (m) => (m === '<sheetData/>' ? `<sheetData>${empty(r)}</sheetData>` : `${empty(r)}</sheetData>`))
  }
  return out
}

const SANDBOX = join(homedir(), 'Library/Containers/com.microsoft.Excel/Data/tmp')
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

/**
 * A case's book in Excel's sandbox, the AppleScript `script` over it (arguments: the path, the rectangles as lines
 * `sheet<TAB>range<TAB>mirror range`, then `more`): its output. A window that is not a book stops the run and is
 * left open (nothing is clicked); so does Excel not answering in 120 s.
 */
export async function runOnBook(id: string, script: string, book: BookInput, read: RangeRef[], more: string[], parts?: FileParts): Promise<string> {
  const before = dialogText()
  if (before) throw new Error(`${id}: Excel has a window that is not a book: ${before}`)
  const dir = join(SANDBOX, `fcase-${process.pid}`)
  mkdirSync(dir, { recursive: true })
  try {
    const file = join(dir, 'book.xlsx')
    await writeCaseBook(book, read, file, parts)
    const at = mirrorRanges(read)
    const child = spawn('osascript', [script, file, read.map((g, i) => `${g.sheet}\t${g.range}\t${at[i]}`).join('\n'), ...more])
    let out = ''
    let err = ''
    child.stdout.on('data', (d: Buffer) => (out += d.toString('utf8')))
    child.stderr.on('data', (d: Buffer) => (err += d.toString('utf8')))
    const exited = new Promise<number | null>((res) => child.on('close', res))
    const until = Date.now() + 120_000
    for (;;) {
      const code = await Promise.race([exited, sleep(1000).then(() => undefined)])
      if (code !== undefined) {
        const after = dialogText()
        if (after) throw new Error(`${id}: Excel window after the case: ${after}`)
        if (code !== 0) throw new Error(`${id}: osascript ${code}: ${err.trim().split('\n')[0]}`)
        return out
      }
      if (excelDialogs().length) {
        child.kill()
        throw new Error(`${id}: Excel shows a window, stopped without clicking: ${dialogText()}`)
      }
      if (Date.now() > until) {
        child.kill()
        throw new Error(`${id}: Excel did not answer in 120 s`)
      }
    }
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
}

/** A range's `value` or `formula` as AppleScript returns it (a cell alone is not a list) as rows. */
const grid = (x: unknown): unknown[][] => (Array.isArray(x) ? (Array.isArray(x[0]) ? (x as unknown[][]) : [x]) : [[x]])

/** The answer from the read of each rectangle: its mirrors' types, numbers and texts, and its formulas. */
export function answerOf(read: RangeRef[], raw: [unknown, unknown, unknown, unknown][]): RectsAnswer {
  const out: RectsAnswer = { values: {}, formulas: {} }
  read.forEach((g, i) => {
    const [t, nums, strs, fs] = raw[i].map(grid)
    const { cols, addrs } = cellsOf(g.range)
    addrs.forEach((a, k) => {
      const [r, j] = [Math.floor(k / cols), k % cols]
      const key = `${g.sheet}!${a}`
      const type = Number(t[r][j])
      const f = String(fs[r][j])
      if (f.startsWith('=')) out.formulas[key] = f
      if (type === -5) return
      out.values[key] =
        type === 0 ? { t: 'n', v: Number(nums[r][j]) } : type === -1 ? { t: 'b', v: true } : type === -2 ? { t: 'b', v: false } : type === -3 ? { t: 's', v: String(strs[r][j]) } : type > 0 ? { t: 'e', v: ERROR_TYPE[type] ?? 'ERROR' } : { t: 'other', v: `type ${type}` }
    })
  })
  return out
}
