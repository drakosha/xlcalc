/**
 * The cells of a worksheet's `<sheetData>`: values by type (shared, inline and formula strings, numbers, booleans,
 * errors, ISO dates of Strict files), formulas with their stored values, rows' heights, styles, hidden rows and outline levels. A cell
 * keeps of its look what its row or column does not give it (`ownLook`). A cell is placed by its `r`, or after the
 * previous one when it has none (rows likewise). Shared formulas and array ranges
 * are collected for the sheet reader to resolve once every cell is known.
 */
import type { CellInput, CellValue, ColumnInput, ErrorCode, NumberFormat, RowInput, Scalar } from '../../contract.js'
import { lookOf, type Look } from '../../layout/lines.js'
import { entryAt, type Line } from '../../layout/spans.js'
import { colToIndex, dateToSerial, formatAddress, MAX_COLS, MAX_ROWS } from '../../model.js'
import { CELL_ERRORS, ERROR_CODE, FORMULA_ERRORS } from '../../syntax/errors.js'
import type { Issues, Where } from '../issues.js'
import type { XlsxCachedCell } from '../types.js'
import type { Xf } from '../styles/read.js'
import { attributes, flag, ownEntry } from '../xml/scan.js'
import { readText } from '../xml/text.js'
import type { CellMetadata } from './metadata.js'
import { CellsLayer } from '../../looks/layer.js'
import { hasLook } from '../../looks/sheet.js'
import { stringItem, type Strings } from './strings.js'

export interface CellsContext {
  sheet: string
  strings: Strings
  xfs: Xf[]
  meta: CellMetadata
  /** The day 0 of the book's date system in the 1900 system (`epochOf`): an ISO date's serial counts from it. */
  epoch: number
  issues: Issues
  /** The book's style with nothing set, added to its styles when first asked for; whether a style is it. */
  plain(): number
  isPlain(style: number): boolean
}

/** A shared formula: its master cell and text, and the cells that take it. */
export interface Shared {
  master?: string
  text?: string
  copies: string[]
}

export interface ArrayFormula {
  anchor: string
  ref: string
  dynamic: boolean
}

export interface SheetCells {
  /** The cells with content (a value, a formula of any kind, an error), with their own looks. */
  cells: Record<string, CellInput>
  /** The looks of the cells without content, which take no entry (`SheetInput.looks`). */
  looks: CellsLayer
  cached: Record<string, XlsxCachedCell>
  rows: Map<number, RowInput>
  shared: Map<string, Shared>
  arrays: ArrayFormula[]
  /** Cells with formulas of the old kind, shared ones among them: not array formulas. */
  old: Set<string>
  /** Cells holding an error as their value, read as a formula of it (unless an array spilled it there). */
  errors: Set<string>
  /** The cells given a formula, shared formulas' masters among them, in the file's order. */
  formulas: string[]
  /**
   * Whether a cell comes at or before one given before it, as a second `<c>` of an address does, or lies outside the
   * formatted row it is given in, which takes its look: a cell with content may then lie under a look.
   */
  unordered: boolean
}

/** The model's codes of the errors a cell may hold. */
export const ERRORS: Record<string, ErrorCode> = Object.fromEntries(CELL_ERRORS.map((text) => [text, ERROR_CODE[text]]))

/** Errors the engine reads in a formula: what an error stored as a cell's value is read as. */
const LITERAL = new Set(FORMULA_ERRORS)

// a tag holds no `<`, nor a value, formula or inline string the opening tag of its name: an element left open fails at
// the next tag, where searching the rest of the cell for its closing tag made a cell of them quadratic
const TAG = /<([\w.-]+:)?(row|c)((?:\s[^<>]*?)?)(\/?)>/g
const V = /<(?:[\w.-]+:)?v(?:\s[^<>]*?)?(?:\/>|>((?:[^<]|<(?!(?:[\w.-]+:)?v[\s/>]))*?)<\/(?:[\w.-]+:)?v>)/
const F = /<(?:[\w.-]+:)?f((?:\s[^<>]*?)?)(?:\/>|>((?:[^<]|<(?!(?:[\w.-]+:)?f[\s/>]))*?)<\/(?:[\w.-]+:)?f>)/
const IS = /<(?:[\w.-]+:)?is(?:\s[^<>]*)?>((?:[^<]|<(?!(?:[\w.-]+:)?is[\s/>]))*?)<\/(?:[\w.-]+:)?is>/
const ADDR = /^([A-Za-z]{1,3})([1-9][0-9]*)$/
// a row number and a shared string's index, whole numbers (spaces about them as XML collapses them); a number's value
const ROW = /^\s*\+?\d+\s*$/
const INDEX = /^\s*\d+\s*$/
const DECIMAL = /^\s*[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?\s*$/
const BOOLEANS = new Map([['1', true], ['true', true], ['0', false], ['false', false]])
// the opening tags of a cell, a value, a formula and an inline string, which a body holds when one was left open
const OPEN_C = /<(?:[\w.-]+:)?c[\s/>]/
const OPEN_V = /<(?:[\w.-]+:)?v[\s/>]/
const OPEN_F = /<(?:[\w.-]+:)?f[\s/>]/
const OPEN_IS = /<(?:[\w.-]+:)?is[\s/>]/

/** A number in decimal notation (xsd:double: no hex, no infinity), NaN for any other text, an empty one too. */
export const decimal = (raw: string): number => (DECIMAL.test(raw) ? Number(raw) : NaN)


/**
 * The cells of `<sheetData>`; `cols`, the sheet's columns. A row with a cell format of its own is read with its cells:
 * Excel shows a cell of it without a `<c>` by the row's whole format, so where such a cell lies under a column with a
 * style or a format the row has not, the row gets the plain style or General (Excel and the writer write those cells).
 */
export function readCells(data: string, ctx: CellsContext, cols: Line<ColumnInput>[] = []): SheetCells {
  const out: SheetCells = { cells: {}, looks: new CellsLayer(), cached: {}, rows: new Map(), shared: new Map(), arrays: [], old: new Set(), errors: new Set(), formulas: [], unordered: false }
  const { looks } = out
  const { issues, sheet } = ctx
  const looked = cols.filter(([, , e]) => hasLook(e))
  const column = (c: number) => (looked.length ? entryAt(looked, c) : undefined)
  let row = -1
  let col = -1
  // the place of the last cell in order
  let [lastRow, lastCol] = [-1, -1]
  // the current row, and when it has a cell format of its own its cells with their own looks, until it ends
  let line: RowInput | undefined
  let lineRow = -1
  let formatted: { addr: string; col: number; own: Look; cell: CellInput }[] | undefined
  const endRow = () => {
    if (!formatted || !line) return
    const at = new Set(formatted.map((x) => x.col))
    const missing = (lo: number, hi: number) => [...at].filter((c) => c >= lo && c <= hi).length < hi - lo + 1
    if (line.style === undefined && looked.some(([lo, hi, e]) => e.style !== undefined && missing(lo, hi))) line.style = ctx.plain()
    if (line.numberFormat === undefined && looked.some(([lo, hi, e]) => e.numberFormat !== undefined && missing(lo, hi))) line.numberFormat = 'General'
    if (Object.keys(line).length) out.rows.set(lineRow + 1, line)
    for (const x of formatted) {
      const look = ownLook(x.own, lookOf(undefined, line, column(x.col)), ctx)
      if (Object.keys(x.cell).length) out.cells[x.addr] = { ...look, ...x.cell }
      else {
        // an address given twice keeps the content another <c> gave it, as in a row without a format
        if (out.cells[x.addr] === x.cell) delete out.cells[x.addr]
        looks.add(lineRow, x.col, looks.looks.of(look))
      }
    }
    formatted = undefined
  }
  TAG.lastIndex = 0
  let m: RegExpExecArray | null
  while ((m = TAG.exec(data))) {
    const a = attributes(m[3])
    if (m[2] === 'row') {
      endRow()
      // a row number that is no whole number places the row after the previous one, as a row without one
      const numbered = !!a.r && ROW.test(a.r)
      row = numbered ? Number(a.r) - 1 : row + 1
      if (a.r && !numbered) issues.add('part', { key: 'rowNumberBad' }, { sheet, cell: `${row + 1}:${row + 1}` })
      col = -1
      line = {}
      lineRow = row
      // a height Excel fitted has no customHeight; Excel for Mac writes its highest, 409.5, as 409.6
      if (a.ht !== undefined) {
        const height = decimal(a.ht)
        if (Number.isFinite(height) && height >= 0) Object.assign(line, { height: Math.min(409.5, height) }, flag(a.customHeight) ? {} : { bestFit: true })
        else issues.add('part', { key: 'heightBad' }, { sheet, cell: `${row + 1}:${row + 1}` })
      }
      // Excel hides a row of height 0 whatever its flag
      if (flag(a.hidden) || line.height === 0) line.hidden = true
      if (a.s !== undefined && flag(a.customFormat)) {
        const where = { sheet, cell: `${row + 1}:${row + 1}` }
        Object.assign(line, look(xfAt(a.s, ctx, where), issues, where))
        formatted = []
      }
      Object.assign(line, outlineOf(a))
      if (row < 0 || row >= MAX_ROWS) {
        // a row outside the sheet: its cells are reported as they come, what it gives them here
        if (formatted || Object.keys(line).length) issues.add('part', { key: 'cellOutside' }, { sheet, cell: `${row + 1}:${row + 1}` })
        line = undefined
        formatted = undefined
        continue
      }
      if (!formatted && Object.keys(line).length) out.rows.set(row + 1, line)
      continue
    }
    const ref = a.r && ADDR.exec(a.r)
    if (ref) {
      col = colToIndex(ref[1])
      row = Number(ref[2]) - 1
    } else col++
    let body = ''
    // an element of the cell left open: the cell, or one it holds, runs into what follows
    let open = false
    if (!m[4]) {
      const close = `</${m[1] ?? ''}c>`
      const end = data.indexOf(close, TAG.lastIndex)
      body = data.slice(TAG.lastIndex, end < 0 ? data.length : end)
      TAG.lastIndex = end < 0 ? data.length : end + close.length
      open = end < 0 || (body.includes(`<${m[1] ?? ''}c`) && OPEN_C.test(body))
    }
    if (row < 0 || row >= MAX_ROWS || col >= MAX_COLS) {
      issues.add('part', { key: 'cellOutside' }, { sheet, cell: a.r ?? `${row + 1}:${col + 1}` })
      continue
    }
    if (row < lastRow || (row === lastRow && col <= lastCol) || (formatted && row !== lineRow)) out.unordered = true
    else {
      lastRow = row
      lastCol = col
    }
    // the address as written when it is already the model's spelling (upper case), else spelled out
    const addr = ref && ref[1] === ref[1].toUpperCase() ? a.r! : formatAddress(row, col)
    const xf = a.s === undefined ? undefined : xfAt(a.s, ctx, { sheet, cell: addr })
    const own = look(xf, issues, { sheet, cell: addr })
    const cell: CellInput = formatted ? {} : ownLook(own, lookOf(undefined, line, column(col)), ctx)
    // a cell of a formatted row takes its look when the row ends
    formatted?.push({ addr, col, own, cell })
    // value metadata: the type of an error Excel stores as #VALUE! (#SPILL!, #CALC!), or a rich value the model does not have
    const rich = a.vm === undefined ? undefined : ctx.meta.error(Number(a.vm))
    if (a.vm !== undefined && rich === undefined) issues.add('cellMetadata', { key: 'valueMetadata' }, { sheet, cell: addr })
    const t = a.t ?? 'n'
    // Excel writes an empty text result as `<v/>`
    const v = body.includes('v') ? V.exec(body) : null
    if (!v && body.includes('v') && OPEN_V.test(body)) open = true
    const raw = v ? (v[1] ?? '') : undefined
    let value: Scalar | undefined
    if (t === 's') {
      if (raw !== undefined) {
        const k = INDEX.test(raw) ? Number(raw) : -1
        value = ctx.strings.text[k]
        if (value === undefined) issues.add('part', { key: 'stringIndexBad' }, { sheet, cell: addr })
        else if (ctx.strings.rich.has(k)) issues.add('richText', { key: 'richText' }, { sheet, cell: addr })
      }
    } else if (t === 'inlineStr') {
      const is = IS.exec(body)
      if (!is && OPEN_IS.test(body)) open = true
      if (is) {
        const s = stringItem(is[1])
        value = s.text
        if (s.rich) issues.add('richText', { key: 'richText' }, { sheet, cell: addr })
      }
    } else if (raw !== undefined && t !== 'e') {
      if (t === 'str') value = readText(raw)
      else if (t === 'd') value = isoDate(raw, ctx, addr)
      else if (raw !== '' || t === 'b') {
        // a number in decimal notation, a boolean as 1, 0, true or false
        const v = t === 'b' ? BOOLEANS.get(raw.trim()) : decimal(raw)
        if (v === undefined || (typeof v === 'number' && !Number.isFinite(v))) issues.add('part', { key: 'valueBad' }, { sheet, cell: addr })
        else value = v
      }
    }
    const errText = t === 'e' && raw !== undefined ? (rich ?? readText(raw).trim()) : undefined
    const f = body.includes('f') ? F.exec(body) : null
    if (!f && body.includes('f') && OPEN_F.test(body)) open = true
    if (open) issues.add('part', { key: 'cellOpen' }, { sheet, cell: addr })
    const fa = f ? attributes(f[1]) : {}
    const text = f?.[2] === undefined ? '' : readText(f[2])
    // `<f ca="1"/>` with no text only flags a value Excel recalculates (a cell a volatile array spilled into)
    if (f && fa.t !== 'dataTable' && (text !== '' || fa.t === 'shared')) {
      if (errText !== undefined) {
        const code = ownEntry(ERRORS, errText)
        out.cached[addr] = { value: { error: code ?? 'ERROR' } }
        if (!code) issues.add('errorValue', { key: 'errorUnknown' }, { sheet, cell: addr })
      } else if (value !== undefined) out.cached[addr] = { value }
      if (fa.t === 'shared' && fa.si !== undefined) {
        let s = out.shared.get(fa.si)
        if (!s) out.shared.set(fa.si, (s = { copies: [] }))
        if (text && s.text === undefined) {
          s.master = addr
          out.formulas.push(addr)
          s.text = text
        } else s.copies.push(addr)
        out.old.add(addr)
      } else {
        const dynamic = fa.t === 'array' && a.cm !== undefined && ctx.meta.dynamic(Number(a.cm))
        if (fa.t === 'array') out.arrays.push({ anchor: addr, ref: (fa.ref ?? addr).toUpperCase(), dynamic })
        if (fa.t !== 'array') out.old.add(addr)
        cell.formula = text
        out.formulas.push(addr)
      }
    } else {
      if (fa.t === 'dataTable') issues.add('dataTable', { key: 'dataTable' }, { sheet, cell: addr })
      if (errText !== undefined) {
        if (LITERAL.has(errText)) {
          cell.formula = errText
          out.formulas.push(addr)
          out.cached[addr] = { value: { error: ERRORS[errText] } }
          out.errors.add(addr)
        } else issues.add('errorValue', { key: 'errorUnspelled' }, { sheet, cell: addr })
      } else if (value !== undefined) cell.value = value
    }
    // a cell of a look alone goes to the sheet's looks, unless a formula is spread to it (a shared one's copy, `<f>` without text)
    if (formatted || cell.value !== undefined || cell.formula !== undefined || (f && fa.t === 'shared')) out.cells[addr] = cell
    else looks.add(row, col, looks.looks.of(cell))
  }
  endRow()
  return out
}

/** A Strict file's date as its serial in the book's date system; one that is no date (`garbage`, `2025-02-30`) is left out and reported. */
function isoDate(raw: string, ctx: CellsContext, cell: string): number | undefined {
  try {
    return dateToSerial(raw) - ctx.epoch
  } catch {
    ctx.issues.add('part', { key: 'dateInvalid' }, { sheet: ctx.sheet, cell })
    return undefined
  }
}

/** Sets a cell's entry; an address the cells did not have goes to `added`, the cells the reader adds to the file's. */
export function setCell(cells: Record<string, CellInput>, addr: string, cell: CellInput, added?: Set<string>): void {
  if (added && cells[addr] === undefined) added.add(addr)
  cells[addr] = cell
}

/** A row's or a column's outline level (1 to 7) and its collapsed mark, as fields of its entry. */
export function outlineOf(a: Record<string, string>): { outlineLevel?: number; collapsed?: boolean } {
  const level = Math.trunc(Number(a.outlineLevel ?? 0))
  return { ...(level > 0 ? { outlineLevel: Math.min(level, 7) } : {}), ...(flag(a.collapsed) ? { collapsed: true } : {}) }
}

/**
 * A stored cell's style and format over the look its row or column gives it (`line`): what they give is left out, so
 * a book written and read back gives no cell a look it did not have, and a default they would not give is spelled out
 * (the plain style, `General`), as a stored cell shows its own look alone.
 */
function ownLook(own: Look, line: Look, ctx: Pick<CellsContext, 'plain' | 'isPlain'>): CellInput {
  const style = line.style !== undefined && ctx.isPlain(line.style) ? undefined : line.style
  const format = line.numberFormat === 'General' || line.numberFormat === 0 ? undefined : line.numberFormat
  const out: CellInput = {}
  if (own.style !== style) out.style = own.style ?? ctx.plain()
  if (own.numberFormat !== format) out.numberFormat = own.numberFormat ?? 'General'
  return out
}

/**
 * The cell format a cell, a row or a column names by its index; one the styles part does not have is reported (the
 * first of a file without cell formats is the default, which needs none).
 */
export function xfAt(index: string, ctx: Pick<CellsContext, 'xfs' | 'issues'>, where: Where): Xf | undefined {
  const k = Number(index)
  const xf = ctx.xfs[k]
  if (xf === undefined && !(k === 0 && !ctx.xfs.length)) ctx.issues.add('part', { key: 'styleIndexBad' }, where)
  return xf
}

/** The style and number format of a cell format, as fields of a cell, a row or a column. */
export function look(xf: Xf | undefined, issues: Issues, where: Where): { style?: number; numberFormat?: NumberFormat } {
  if (!xf) return {}
  for (const lost of xf.lost) issues.add('style', { key: 'styleLost', part: lost }, where)
  if (xf.eastAsian !== undefined) issues.add('numberFormat', { key: 'eastAsianFormat' }, where)
  if (xf.unknownFormat !== undefined) issues.add('numberFormat', { key: 'unknownFormat' }, where)
  const out: { style?: number; numberFormat?: NumberFormat } = {}
  if (xf.style !== undefined) out.style = xf.style
  if (xf.format !== undefined) out.numberFormat = xf.format
  return out
}
