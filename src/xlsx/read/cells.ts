/**
 * The cells of a worksheet's `<sheetData>`: values by type (shared, inline and formula strings, numbers, booleans,
 * errors, ISO dates of Strict files), formulas with their stored values, rows' heights and styles. A cell is placed
 * by its `r`, or after the previous one when it has none (rows likewise). Shared formulas and array ranges are
 * collected for the sheet reader to resolve once every cell is known.
 */
import type { CellInput, CellValue, ErrorCode, RowInput, Scalar } from '../../contract.js'
import { dateToSerial, formatAddress } from '../../model.js'
import type { Issues } from '../issues.js'
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
  /** A number of a date format moved to the 1900 date system (as it is in a book of that system). */
  date(n: number, xf: Xf | undefined): number
  issues: Issues
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
  /** Cells with formulas of the old kind: no dynamic array metadata. */
  old: Set<string>
}

/** Error texts by the model's codes; `#NULL!` is what the engine calls #ERROR!. */
const ERRORS: Record<string, ErrorCode> = {
  '#NULL!': 'ERROR',
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
const V = /<(?:[\w.-]+:)?v(?:\s[^>]*)?>([\s\S]*?)<\/(?:[\w.-]+:)?v>/
const F = /<(?:[\w.-]+:)?f((?:\s[^>]*?)?)(?:\/>|>([\s\S]*?)<\/(?:[\w.-]+:)?f>)/
const IS = /<(?:[\w.-]+:)?is(?:\s[^>]*)?>([\s\S]*?)<\/(?:[\w.-]+:)?is>/
const ADDR = /^([A-Za-z]{1,3})([1-9][0-9]*)$/

const colOf = (letters: string) => {
  let n = 0
  for (let i = 0; i < letters.length; i++) n = n * 26 + ((letters.charCodeAt(i) | 32) - 96)
  return n - 1
}

export function readCells(data: string, ctx: CellsContext): SheetCells {
  const out: SheetCells = { cells: {}, cached: {}, rows: new Map(), shared: new Map(), arrays: [], old: new Set() }
  const { issues, sheet } = ctx
  let row = -1
  let col = -1
  TAG.lastIndex = 0
  let m: RegExpExecArray | null
  while ((m = TAG.exec(data))) {
    const a = attributes(m[3])
    if (m[2] === 'row') {
      row = a.r ? Number(a.r) - 1 : row + 1
      col = -1
      const line: RowInput = {}
      if (flag(a.hidden)) {
        line.height = 0
        issues.add('hidden', 'hidden rows and columns are read with a size of 0', { sheet, cell: `${row + 1}:${row + 1}` })
      } else if (a.ht !== undefined && flag(a.customHeight)) line.height = Number(a.ht)
      if (a.s !== undefined && flag(a.customFormat)) Object.assign(line, look(ctx.xfs[Number(a.s)]))
      if (Number(a.outlineLevel ?? 0) > 0) issues.add('outline', 'grouped rows and columns are not in the model', { sheet, cell: `${row + 1}:${row + 1}` })
      if (Object.keys(line).length) out.rows.set(row + 1, line)
      continue
    }
    const ref = a.r && ADDR.exec(a.r)
    if (ref) {
      col = colOf(ref[1])
      row = Number(ref[2]) - 1
    } else col++
    const addr = formatAddress(row, col)
    let body = ''
    if (!m[4]) {
      const close = `</${m[1] ?? ''}c>`
      const end = data.indexOf(close, TAG.lastIndex)
      body = data.slice(TAG.lastIndex, end < 0 ? data.length : end)
      TAG.lastIndex = end < 0 ? data.length : end + close.length
    }
    const xf = a.s === undefined ? undefined : ctx.xfs[Number(a.s)]
    const cell: CellInput = look(xf)
    if (xf) for (const lost of xf.lost) issues.add('style', `a style's ${lost} is not in the model`, { sheet, cell: addr })
    if (xf?.regional !== undefined) issues.add('numberFormat', 'a built-in date, time or currency format Excel shows by the regional settings, read as its en-US code', { sheet, cell: addr })
    if (xf?.unknownFormat !== undefined) issues.add('numberFormat', 'a built-in number format this reader does not know, left out', { sheet, cell: addr })
    // value metadata: the type of an error Excel stores as #VALUE! (#SPILL!, #CALC!), or a rich value the model does not have
    const rich = a.vm === undefined ? undefined : ctx.meta.error(Number(a.vm))
    if (a.vm !== undefined && rich === undefined) issues.add('cellMetadata', "a cell's value metadata (a data type, a picture in the cell) is not in the model", { sheet, cell: addr })
    const t = a.t ?? 'n'
    const raw = body.includes('v') ? V.exec(body)?.[1] : undefined
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
      else if (t === 'd') value = dateToSerial(raw)
      else if (raw !== '' && Number.isFinite(Number(raw))) value = ctx.date(Number(raw), xf)
    }
    const errText = t === 'e' && raw !== undefined ? (rich ?? readText(raw).trim()) : undefined
    const f = body.includes('f') ? F.exec(body) : null
    const fa = f ? attributes(f[1]) : {}
    if (f && fa.t !== 'dataTable') {
      const text = f[2] === undefined ? '' : readText(f[2])
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
        if (!dynamic) out.old.add(addr)
        cell.formula = text
      }
    } else {
      if (f) issues.add('dataTable', 'a what-if data table (TABLE) is read as its stored values', { sheet, cell: addr })
      if (errText !== undefined) {
        if (LITERAL.has(errText)) {
          cell.formula = errText
          out.cached[addr] = { error: ERRORS[errText] }
          issues.add('errorConstant', 'an error typed into a cell is read as a formula of that error', { sheet, cell: addr })
        } else issues.add('errorValue', 'an error the model has no code for, left out', { sheet, cell: addr })
      } else if (value !== undefined) cell.value = value
    }
    if (cell.value !== undefined || cell.formula !== undefined || cell.style !== undefined || cell.format !== undefined) out.cells[addr] = cell
  }
  return out
}

/** The style and number format of a cell format, as fields of a cell, a row or a column. */
export function look(xf: Xf | undefined): { style?: number; format?: string } {
  if (!xf) return {}
  const out: { style?: number; format?: string } = {}
  if (xf.style !== undefined) out.style = xf.style
  if (xf.format !== undefined) out.format = xf.format
  return out
}
