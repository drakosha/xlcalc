/**
 * A worksheet part from a sheet of the book. Formulas are written as dynamic array formulas (`cm` with the XLDAPR
 * metadata, `t="array"` over the range their array spilled into), as Excel 365 saves them: the engine computes
 * every formula by the rules of dynamic arrays, and so does Excel then. An array formula of the old kind
 * (`CellInput.array`) is written over its area without the metadata, as Excel writes one entered with
 * Ctrl+Shift+Enter. Values come from the result: a formula's value, and the cells its array spilled into. A formula
 * the file cannot hold (it does not parse, or names a sheet the book does not have) is written as its value.
 */
import type { CellInput, CellResult, CellValue, SheetInput } from '../../contract.js'
import { lines } from '../../layout/spans.js'
import { dateToSerial, formatAddress, indexToCol, MAX_COLS, parseAddress } from '../../model.js'
import { parseRange } from '../../structure/area.js'
import { ERROR_TEXT, FormulaError, type Ast } from '../../syntax/ast.js'
import { parse } from '../../syntax/parse.js'
import { sheetKey, walk } from '../../syntax/print.js'
import { toFile } from '../formulas/to-file.js'
import type { Issues } from '../issues.js'
import { escapeAttr, escapeText } from '../xml/text.js'
import type { StyleWriter } from './styles.js'

/** Excel's limit on a cell's text. */
export const MAX_TEXT = 32_767

/**
 * The default width in xlsx units (8 digits and the padding, as Excel for Mac writes it): a column that has a look
 * but no width of its own needs one, as Excel reads a column without a width as hidden.
 */
const DEFAULT_WIDTH = '8.83203125'

export interface SheetContext {
  styles: StyleWriter
  /** The index of a string in the shared strings table. */
  string(s: string): number
  issues: Issues
  /** Upper-case names (`sheetKey`) of the book's sheets. */
  sheets: Set<string>
  mode: 'formulas' | 'values'
  /** The sheet's cells of the result, when there is one. */
  result?: Record<string, CellResult>
  /** The value metadata index (`vm`, 1-based) of an error Excel stores as a rich value: #SPILL!, #CALC!. */
  richError(code: 'SPILL' | 'CALC'): number
}

export interface SheetWritten {
  xml: string
  /** Whether a formula was written (the book then needs the dynamic array metadata). */
  formulas: boolean
}

/**
 * A value as the cell's attributes (`t`, and `vm` for an error Excel keeps as a rich value: the cell holds #VALUE!
 * for older readers) and the content of `<v>`; `shared` writes text through the shared strings.
 */
function valueXml(v: CellValue, ctx: SheetContext, where: { sheet: string; cell: string }, shared: boolean): { t: string; v: string } {
  if (v === null) return { t: '', v: '0' }
  if (typeof v === 'number') return Number.isFinite(v) ? { t: '', v: String(v) } : { t: ' t="e"', v: '#NUM!' }
  if (typeof v === 'boolean') return { t: ' t="b"', v: v ? '1' : '0' }
  if (typeof v === 'string') {
    let s = v
    if (s.length > MAX_TEXT) {
      s = s.slice(0, MAX_TEXT)
      ctx.issues.add('textLength', "text over Excel's 32,767 characters is cut", where)
    }
    return shared ? { t: ' t="s"', v: String(ctx.string(s)) } : { t: ' t="str"', v: escapeText(s) }
  }
  if (v.error === 'CYCLE') {
    ctx.issues.add('errorValue', 'a circular reference (#CYCLE!) is written as 0, as Excel stores it', where)
    return { t: '', v: '0' }
  }
  if (v.error === 'SPILL' || v.error === 'CALC') return { t: ` t="e" vm="${ctx.richError(v.error)}"`, v: '#VALUE!' }
  return { t: ' t="e"', v: v.error === 'ERROR' ? '#NULL!' : ERROR_TEXT[v.error] }
}

/** The formula's tree when the file can hold it: it parses and names only the book's sheets. */
function fileTree(formula: string, sheets: Set<string>): Ast | null {
  let ast: Ast
  try {
    ast = parse(formula)
  } catch (e) {
    if (e instanceof FormulaError) return null
    throw e
  }
  let ok = true
  const known = (sheet: string | null | undefined) => sheet == null || sheets.has(sheetKey(sheet))
  walk(ast, (n) => {
    if (n.k === 'ref3d' ? !known(n.first) || !known(n.last) : (n.k === 'ref' || n.k === 'err' || n.k === 'name') && !known(n.sheet)) ok = false
  })
  return ok ? ast : null
}

/** The cells of a rectangle `A1:C9`, row by row. */
function* rectCells(range: string): Generator<string> {
  const r = parseRange(range)!
  for (let row = r.r1; row <= r.r2; row++) for (let c = r.c1; c <= r.c2; c++) yield formatAddress(row, c)
}

export function writeSheet(sheet: SheetInput, index: number, ctx: SheetContext): SheetWritten {
  const { styles, issues, mode, result } = ctx
  const out = new Map<number, string>()
  const rowsUsed = new Set<number>()
  let formulas = false
  const put = (addr: string, xml: string) => {
    const { row, col } = parseAddress(addr)
    out.set(row * MAX_COLS + col, xml)
    rowsUsed.add(row)
  }
  const where = (cell: string) => ({ sheet: sheet.name, cell })
  const sAttr = (c: CellInput | undefined) => {
    const xf = c ? styles.xf(c.style, c.format) : 0
    return xf ? ` s="${xf}"` : ''
  }
  // `cm` marks a dynamic array formula
  const cellXml = (addr: string, c: CellInput | undefined, v: CellValue | undefined, f = '', dynamic = f !== '') => {
    const cm = dynamic ? ' cm="1"' : ''
    if (v === undefined) return `<c r="${addr}"${sAttr(c)}${cm}${f ? `>${f}</c>` : '/>'}`
    const x = valueXml(v, ctx, where(addr), !f)
    return `<c r="${addr}"${sAttr(c)}${x.t}${cm}>${f}<v>${x.v}</v></c>`
  }
  for (const [addr, c] of Object.entries(sheet.cells)) {
    if (c.meta !== undefined) issues.add('meta', 'cell meta has no place in a file', where(addr))
    const res = result?.[addr]
    if (c.formula === undefined) {
      const v = c.value === undefined || c.value === null ? undefined : typeof c.value === 'object' ? dateToSerial(c.value.date) : c.value
      put(addr, cellXml(addr, c, v))
      continue
    }
    const tree = mode === 'formulas' ? fileTree(c.formula, ctx.sheets) : null
    if (mode === 'formulas' && !tree) issues.add('formula', 'a formula the file cannot hold (it does not parse, or names a sheet the book does not have) is written as its value', where(addr))
    if (!tree) {
      put(addr, cellXml(addr, c, res?.value))
      continue
    }
    formulas = true
    const ref = c.array ?? res?.spill ?? addr
    put(addr, cellXml(addr, c, res?.value, `<f t="array" ref="${ref}">${escapeText(toFile(tree))}</f>`, c.array === undefined))
  }
  // the cells arrays spilled into, with the look their own cells give them
  if (result)
    for (const [addr, r] of Object.entries(result)) {
      if (r.spillOf !== undefined) put(addr, cellXml(addr, sheet.cells[addr], r.value))
      if (r.spill !== undefined && r.spillFill !== undefined && mode === 'values')
        for (const a of rectCells(r.spill)) if (!result[a]) put(a, cellXml(a, sheet.cells[a], r.spillFill))
    }
  const rowLines = lines(sheet.rows, 'r')
  for (const [lo, hi] of rowLines) for (let r = lo; r <= hi; r++) rowsUsed.add(r)
  const rowAttrs = new Map<number, string>()
  for (const [lo, hi, e] of rowLines) {
    const xf = e.style !== undefined || e.format !== undefined ? styles.xf(e.style, e.format) : 0
    const attrs = (e.height === 0 ? ' hidden="1"' : e.height !== undefined ? ` ht="${e.height}" customHeight="1"` : '') + (xf ? ` s="${xf}" customFormat="1"` : '')
    for (let r = lo; r <= hi; r++) rowAttrs.set(r, attrs)
  }
  const keys = [...out.keys()].sort((a, b) => a - b)
  const data: string[] = []
  let k = 0
  for (const row of [...rowsUsed].sort((a, b) => a - b)) {
    const cells: string[] = []
    while (k < keys.length && Math.floor(keys[k] / MAX_COLS) === row) cells.push(out.get(keys[k++])!)
    data.push(cells.length ? `<row r="${row + 1}"${rowAttrs.get(row) ?? ''}>${cells.join('')}</row>` : `<row r="${row + 1}"${rowAttrs.get(row) ?? ''}/>`)
  }
  const first = keys.length ? keys[0] : 0
  const last = keys.length ? keys[keys.length - 1] : 0
  const minCol = keys.reduce((m, x) => Math.min(m, x % MAX_COLS), MAX_COLS)
  const maxCol = keys.reduce((m, x) => Math.max(m, x % MAX_COLS), 0)
  const dim = keys.length ? `${indexToCol(minCol)}${Math.floor(first / MAX_COLS) + 1}:${indexToCol(maxCol)}${Math.floor(last / MAX_COLS) + 1}` : 'A1'
  const cols = lines(sheet.columns, 'c').map(([lo, hi, e]) => {
    const xf = e.style !== undefined || e.format !== undefined ? styles.xf(e.style, e.format) : 0
    const size = e.width === 0 ? ' width="0" customWidth="1" hidden="1"' : e.width !== undefined ? ` width="${e.width}" customWidth="1"` : ` width="${DEFAULT_WIDTH}"`
    return `<col min="${lo + 1}" max="${hi + 1}"${size}${xf ? ` style="${xf}"` : ''}/>`
  })
  const merges = sheet.merges ?? []
  const xml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    `<dimension ref="${dim}"/><sheetViews><sheetView workbookViewId="0"${index === 0 ? ' tabSelected="1"' : ''}/></sheetViews>` +
    // without it Excel for Mac takes a standard width of 10 digits, not the 8 the model assumes
    '<sheetFormatPr defaultRowHeight="15"/>' +
    (cols.length ? `<cols>${cols.join('')}</cols>` : '') +
    `<sheetData>${data.join('')}</sheetData>` +
    (merges.length ? `<mergeCells count="${merges.length}">${merges.map((m) => `<mergeCell ref="${escapeAttr(m)}"/>`).join('')}</mergeCells>` : '') +
    '</worksheet>'
  return { xml, formulas }
}
