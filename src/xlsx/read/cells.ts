/**
 * The cells of a worksheet's `<sheetData>`: values by type (shared, inline and formula strings, numbers, booleans,
 * errors, ISO dates of Strict files), formulas with their stored values, rows' heights, styles, hidden rows and outline levels. A cell
 * keeps of its look what its row or column does not give it (`ownLook`). A cell is placed by its `r`, or after the
 * previous one when it has none (rows likewise). Shared formulas and array ranges
 * are collected for the sheet reader to resolve once every cell is known.
 */
import type { CellInput, CellValue, ErrorCode, NumberFormat, RowInput, Scalar } from '../../contract.js'
import { lookOf, type Look } from '../../layout/lines.js'
import { dateToSerial, formatAddress, MAX_COLS, MAX_ROWS } from '../../model.js'
import type { Issues, Where } from '../issues.js'
import type { Xf } from '../styles/read.js'
import { attributes, flag } from '../xml/scan.js'
import { readText } from '../xml/text.js'
import type { CellMetadata } from './metadata.js'
import { stringItem, type Strings } from './strings.js'

export interface CellsContext {
  sheet: string
  strings: Strings
  xfs: Xf[]
  meta: CellMetadata
  /** The day 0 of the book's date system in the 1900 system (`epochOf`): an ISO date's serial counts from it. */
  epoch: number
  issues: Issues
  /** The book's style with nothing set, added to its styles when first asked for. */
  plain(): number
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
  cells: Record<string, CellInput>
  cached: Record<string, CellValue>
  rows: Map<number, RowInput>
  shared: Map<string, Shared>
  arrays: ArrayFormula[]
  /** Cells with formulas of the old kind, shared ones among them: not array formulas. */
  old: Set<string>
  /** Cells holding an error as their value, read as a formula of it (unless an array spilled it there). */
  errors: Set<string>
}

/** Error texts by the model's codes. */
const ERRORS: Record<string, ErrorCode> = {
  '#NULL!': 'NULL',
  '#DIV/0!': 'DIV0',
  '#VALUE!': 'VALUE',
  '#REF!': 'REF',
  '#NAME?': 'NAME',
  '#NUM!': 'NUM',
  '#N/A': 'NA',
  '#SPILL!': 'SPILL',
  '#CALC!': 'CALC',
}

/** Errors the engine reads in a formula: what an error stored as a cell's value is read as. */
const LITERAL = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#SPILL!'])

const TAG = /<([\w.-]+:)?(row|c)((?:\s[^>]*?)?)(\/?)>/g
const V = /<(?:[\w.-]+:)?v(?:\s[^>]*?)?(?:\/>|>([\s\S]*?)<\/(?:[\w.-]+:)?v>)/
const F = /<(?:[\w.-]+:)?f((?:\s[^>]*?)?)(?:\/>|>([\s\S]*?)<\/(?:[\w.-]+:)?f>)/
const IS = /<(?:[\w.-]+:)?is(?:\s[^>]*)?>([\s\S]*?)<\/(?:[\w.-]+:)?is>/
const ADDR = /^([A-Za-z]{1,3})([1-9][0-9]*)$/

const colOf = (letters: string) => {
  let n = 0
  for (let i = 0; i < letters.length; i++) n = n * 26 + ((letters.charCodeAt(i) | 32) - 96)
  return n - 1
}

/** The cells of `<sheetData>`; `column`, the look of a column of the sheet (zero-based). */
export function readCells(data: string, ctx: CellsContext, column: (c: number) => Look | undefined = () => undefined): SheetCells {
  const out: SheetCells = { cells: {}, cached: {}, rows: new Map(), shared: new Map(), arrays: [], old: new Set(), errors: new Set() }
  const { issues, sheet } = ctx
  let row = -1
  let col = -1
  let rowLine: RowInput | undefined
  TAG.lastIndex = 0
  let m: RegExpExecArray | null
  while ((m = TAG.exec(data))) {
    const a = attributes(m[3])
    if (m[2] === 'row') {
      row = a.r ? Number(a.r) - 1 : row + 1
      col = -1
      const line: RowInput = {}
      if (a.ht !== undefined && flag(a.customHeight)) line.height = Number(a.ht)
      // Excel hides a row of height 0 whatever its flag
      if (flag(a.hidden) || line.height === 0) line.hidden = true
      if (a.s !== undefined && flag(a.customFormat)) Object.assign(line, look(ctx.xfs[Number(a.s)], issues, { sheet, cell: `${row + 1}:${row + 1}` }))
      Object.assign(line, outlineOf(a))
      rowLine = line
      if (Object.keys(line).length) out.rows.set(row + 1, line)
      continue
    }
    const ref = a.r && ADDR.exec(a.r)
    if (ref) {
      col = colOf(ref[1])
      row = Number(ref[2]) - 1
    } else col++
    let body = ''
    if (!m[4]) {
      const close = `</${m[1] ?? ''}c>`
      const end = data.indexOf(close, TAG.lastIndex)
      body = data.slice(TAG.lastIndex, end < 0 ? data.length : end)
      TAG.lastIndex = end < 0 ? data.length : end + close.length
    }
    if (row < 0 || row >= MAX_ROWS || col >= MAX_COLS) {
      issues.add('part', 'a cell outside the sheet (past its last row or column) is left out', { sheet, cell: a.r ?? `${row + 1}:${col + 1}` })
      continue
    }
    // the address as written when it is already the model's spelling (upper case), else spelled out
    const addr = ref && ref[1] === ref[1].toUpperCase() ? a.r! : formatAddress(row, col)
    const xf = a.s === undefined ? undefined : ctx.xfs[Number(a.s)]
    const cell: CellInput = ownLook(look(xf, issues, { sheet, cell: addr }), lookOf(undefined, rowLine, column(col)), ctx.plain)
    // value metadata: the type of an error Excel stores as #VALUE! (#SPILL!, #CALC!), or a rich value the model does not have
    const rich = a.vm === undefined ? undefined : ctx.meta.error(Number(a.vm))
    if (a.vm !== undefined && rich === undefined) issues.add('cellMetadata', "a cell's value metadata (a data type, a picture in the cell) is not in the model", { sheet, cell: addr })
    const t = a.t ?? 'n'
    // Excel writes an empty text result as `<v/>`
    const v = body.includes('v') ? V.exec(body) : null
    const raw = v ? (v[1] ?? '') : undefined
    let value: Scalar | undefined
    if (t === 's') {
      if (raw !== undefined) {
        const k = Number(raw)
        value = ctx.strings.text[k] ?? ''
        if (ctx.strings.rich.has(k)) issues.add('richText', 'formatted runs of text are read as plain text', { sheet, cell: addr })
      }
    } else if (t === 'inlineStr') {
      const is = IS.exec(body)
      if (is) {
        const s = stringItem(is[1])
        value = s.text
        if (s.rich) issues.add('richText', 'formatted runs of text are read as plain text', { sheet, cell: addr })
      }
    } else if (raw !== undefined && t !== 'e') {
      if (t === 'str') value = readText(raw)
      else if (t === 'b') value = raw === '1' || raw === 'true'
      else if (t === 'd') value = dateToSerial(raw) - ctx.epoch
      else if (raw !== '' && Number.isFinite(Number(raw))) value = Number(raw)
    }
    const errText = t === 'e' && raw !== undefined ? (rich ?? readText(raw).trim()) : undefined
    const f = body.includes('f') ? F.exec(body) : null
    const fa = f ? attributes(f[1]) : {}
    const text = f?.[2] === undefined ? '' : readText(f[2])
    // `<f ca="1"/>` with no text only flags a value Excel recalculates (a cell a volatile array spilled into)
    if (f && fa.t !== 'dataTable' && (text !== '' || fa.t === 'shared')) {
      if (errText !== undefined) {
        out.cached[addr] = { error: ERRORS[errText] ?? 'ERROR' }
        if (!ERRORS[errText]) issues.add('errorValue', 'a stored error the model has no code for, read as #ERROR!', { sheet, cell: addr })
      } else if (value !== undefined) out.cached[addr] = value
      if (fa.t === 'shared' && fa.si !== undefined) {
        let s = out.shared.get(fa.si)
        if (!s) out.shared.set(fa.si, (s = { copies: [] }))
        if (text && s.text === undefined) {
          s.master = addr
          s.text = text
        } else s.copies.push(addr)
        out.old.add(addr)
      } else {
        const dynamic = fa.t === 'array' && a.cm !== undefined && ctx.meta.dynamic(Number(a.cm))
        if (fa.t === 'array') out.arrays.push({ anchor: addr, ref: (fa.ref ?? addr).toUpperCase(), dynamic })
        if (fa.t !== 'array') out.old.add(addr)
        cell.formula = text
      }
    } else {
      if (fa.t === 'dataTable') issues.add('dataTable', 'a what-if data table (TABLE) is read as its stored values', { sheet, cell: addr })
      if (errText !== undefined) {
        if (LITERAL.has(errText)) {
          cell.formula = errText
          out.cached[addr] = { error: ERRORS[errText] }
          out.errors.add(addr)
        } else issues.add('errorValue', 'an error stored as a value that a formula cannot spell (#CALC!, #GETTING_DATA and the like), left out', { sheet, cell: addr })
      } else if (value !== undefined) cell.value = value
    }
    if (cell.value !== undefined || cell.formula !== undefined || cell.style !== undefined || cell.format !== undefined) out.cells[addr] = cell
  }
  return out
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
function ownLook(own: Look, line: Look, plain: () => number): CellInput {
  const out: CellInput = {}
  if (own.style !== line.style) out.style = own.style ?? plain()
  if (own.format !== line.format) out.format = own.format ?? 'General'
  return out
}

/** The style and number format of a cell format, as fields of a cell, a row or a column. */
export function look(xf: Xf | undefined, issues: Issues, where: Where): { style?: number; format?: NumberFormat } {
  if (!xf) return {}
  for (const lost of xf.lost) issues.add('style', `a style's ${lost} is not in the model`, where)
  if (xf.eastAsian !== undefined) issues.add('numberFormat', 'a built-in East Asian date format the model does not have (50 to 58), read as the code Excel for Mac en-US shows', where)
  if (xf.unknownFormat !== undefined) issues.add('numberFormat', 'a built-in number format this reader does not know, left out', where)
  const out: { style?: number; format?: NumberFormat } = {}
  if (xf.style !== undefined) out.style = xf.style
  if (xf.format !== undefined) out.format = xf.format
  return out
}
