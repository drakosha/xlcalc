/**
 * A worksheet part as a sheet of the book: cells (see `cells.ts`), formulas in the engine's spelling (those of the
 * old kind as the dynamic formulas Excel 365 shows for them, `@` where Excel puts it), shared formulas spread to their
 * cells as a copy and paste spreads them (`restructure`'s copyRange), array formulas of the old kind over their area,
 * the cells an array spilled into left to the stored values, columns, rows, merged areas, the filter mode, notes,
 * hyperlinks and frozen panes (`parts.ts`), and what the sheet has that the model does not, reported.
 */
import type { CellValue, ColumnInput, RowInput, SheetInput } from '../../contract.js'
import { mergeText, overlapping, parseMerge } from '../../layout/merges.js'
import { spanKey } from '../../layout/spans.js'
import { formatAddress, parseAddress, sortAddresses } from '../../model.js'
import { parseRange, type Rect } from '../../structure/area.js'
import { copyRange } from '../../structure/copy.js'
import type { Ast } from '../../syntax/ast.js'
import { canonical } from '../../syntax/print.js'
import type { Issues } from '../issues.js'
import type { XlsxIssueCode } from '../types.js'
import { attributes, element, elements, flag, has, type Element } from '../xml/scan.js'
import { look, readCells, type CellsContext, type SheetCells } from './cells.js'
import { readFormula, reportFormula, type ReadFormula } from './formulas.js'
import type { Rel } from './package.js'
import { readCellParts, readFrozen, type PartsSource } from './parts.js'
import { readTables } from './tables.js'

export interface SheetSource {
  name: string
  xml: string
  rels: Rel[]
  /** The names of the book's sheets: shared formulas may reference any of them. */
  sheets: string[]
  /** A formula of the old kind as the dynamic one (`legacyConverter` over the book's names), undefined when the same. */
  legacy(ast: Ast, sheet: string): Ast | undefined
  /** The sheet's table parts. */
  tables: string[]
  /** The sheet's comments, its VML drawing and threaded comments, and the book's persons, when there are any. */
  parts: Pick<PartsSource, 'comments' | 'vml' | 'threads' | 'persons'>
}

export interface SheetRead {
  sheet: SheetInput
  cached: Record<string, CellValue>
}

/** Elements of a worksheet the model has no place for, outside `<sheetData>`. */
const FEATURES: [local: string, code: XlsxIssueCode, message: string][] = [
  ['conditionalFormatting', 'conditionalFormat', 'conditional formatting is not in the model'],
  ['conditionalFormattings', 'conditionalFormat', 'conditional formatting is not in the model'],
  ['dataValidations', 'dataValidation', 'data validation is not in the model'],
  ['sheetProtection', 'protection', 'sheet protection is not in the model'],
  ['protectedRanges', 'protection', 'sheet protection is not in the model'],
  ['drawing', 'drawing', 'pictures, charts and shapes are not in the model'],
  ['picture', 'drawing', 'pictures, charts and shapes are not in the model'],
  ['oleObjects', 'drawing', 'pictures, charts and shapes are not in the model'],
  ['controls', 'drawing', 'pictures, charts and shapes are not in the model'],
  ['sparklineGroups', 'sparkline', 'sparklines are not in the model'],
  ['headerFooter', 'print', 'print settings (headers and footers, page breaks, page setup) are not in the model'],
  ['rowBreaks', 'print', 'print settings (headers and footers, page breaks, page setup) are not in the model'],
  ['colBreaks', 'print', 'print settings (headers and footers, page breaks, page setup) are not in the model'],
]

/** Parts a sheet links to that the model has no place for, by relationship kind. */
const LINKED: Record<string, [XlsxIssueCode, string]> = {
  pivotTable: ['pivotTable', 'pivot tables are not in the model: their cells are read as stored'],
}

/** The cells of a rectangle that the cells' count makes cheaper to list than the sheet's cells. */
function* cellsOf(r: Rect, keys: () => string[]): Generator<string> {
  if ((r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1) <= 4096) {
    for (let row = r.r1; row <= r.r2; row++) for (let c = r.c1; c <= r.c2; c++) yield formatAddress(row, c)
    return
  }
  for (const a of keys()) {
    const { row, col } = parseAddress(a)
    if (row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2) yield a
  }
}

/** Shared formulas in each of their cells: the master's formula copied and pasted over them. */
function spreadShared(src: SheetSource, read: SheetCells, issues: Issues) {
  for (const s of read.shared.values()) {
    if (s.master === undefined) {
      for (const c of s.copies) issues.add('formula', 'a shared formula without its master cell is left out', { sheet: src.name, cell: c })
      continue
    }
    const formula = read.cells[s.master]!.formula!
    const at = [s.master, ...s.copies].map(parseAddress)
    const box = { r1: Infinity, r2: -1, c1: Infinity, c2: -1 }
    for (const a of at) {
      box.r1 = Math.min(box.r1, a.row)
      box.r2 = Math.max(box.r2, a.row)
      box.c1 = Math.min(box.c1, a.col)
      box.c2 = Math.max(box.c2, a.col)
    }
    // copies spread thin over a large area are pasted one by one
    const targets = (box.r2 - box.r1 + 1) * (box.c2 - box.c1 + 1) > 4 * at.length ? s.copies : [`${formatAddress(box.r1, box.c1)}:${formatAddress(box.r2, box.c2)}`]
    const book = { sheets: src.sheets.map((name) => ({ name, cells: name === src.name ? { [s.master!]: { formula } } : {} })) }
    for (const range of targets) {
      // the pure copy and paste, without restructure's checks and values (a paste from a filtered sheet computes the book)
      const r = copyRange(book, src.name, parseRange(s.master)!, src.name, parseRange(range)!, false)
      const pasted = r.ok ? r.book.sheets.find((x) => x.name === src.name)!.cells : {}
      for (const c of s.copies) {
        const f = pasted[c]?.formula
        if (f !== undefined) read.cells[c] = { ...read.cells[c], formula: f }
      }
    }
  }
}

/** The formulas of the sheet in the engine's spelling, those of the old kind as dynamic ones. */
function convert(src: SheetSource, read: SheetCells, issues: Issues): Map<string, ReadFormula> {
  const out = new Map<string, ReadFormula>()
  for (const addr of Object.keys(read.cells)) {
    const cell = read.cells[addr]
    if (cell.formula === undefined) continue
    const f = readFormula(cell.formula, issues, { sheet: src.name, cell: addr })
    const dynamic = f.ast && read.old.has(addr) ? src.legacy(f.ast, src.name) : undefined
    cell.formula = dynamic ? canonical(dynamic) : f.formula
    out.set(addr, f)
  }
  return out
}

/** The cells an array formula spilled into hold stored values: they go to `cached`, keeping their look. */
function arrays(read: SheetCells) {
  const keys = () => Object.keys(read.cells)
  for (const a of read.arrays) {
    const r = parseRange(a.ref)
    if (!r) continue
    const anchor = read.cells[a.anchor]
    if (!a.dynamic && anchor?.formula !== undefined) anchor.array = a.ref
    for (const addr of cellsOf(r, keys)) {
      const cell = read.cells[addr]
      // an error the array spilled is stored as the cell's value, which the reader took for a formula of it
      const spilled = read.errors.delete(addr)
      if (addr === a.anchor || !cell || (cell.formula !== undefined && !spilled)) continue
      if (cell.value !== undefined) read.cached[addr] = cell.value as CellValue
      const { value: _, formula: __, ...rest } = cell
      if (Object.keys(rest).length) read.cells[addr] = rest
      else delete read.cells[addr]
    }
  }
}

/** Lines (rows or columns) with equal settings next to each other as one span. */
function spans<T>(lines: [lo: number, hi: number, T][], axis: 'r' | 'c'): Record<string, T> | undefined {
  lines.sort((x, y) => x[0] - y[0])
  const out: [number, number, T, string][] = []
  for (const [lo, hi, v] of lines) {
    const key = JSON.stringify(v)
    const last = out.at(-1)
    if (last && last[1] + 1 === lo && last[3] === key) last[1] = hi
    else out.push([lo, hi, v, key])
  }
  return out.length ? Object.fromEntries(out.map(([lo, hi, v]) => [spanKey(lo, hi, axis), v])) : undefined
}

function columns(head: string, read: CellsContext): Record<string, ColumnInput> | undefined {
  const lines: [number, number, ColumnInput][] = []
  for (const e of elements(element(head, 'cols')?.body ?? '', 'col')) {
    const a = attributes(e.attrs)
    const [lo, hi] = [Number(a.min) - 1, Number(a.max) - 1]
    if (!(lo >= 0 && hi >= lo && hi < 16384)) continue
    const col: ColumnInput = {}
    const where = { sheet: read.sheet, cell: `${spanKey(lo, hi, 'c').replace(/^([A-Z]+)$/, '$1:$1')}` }
    if (a.width !== undefined && flag(a.customWidth)) col.width = Math.min(255, Number(a.width))
    // Excel hides a column of width 0 whatever its flag
    if (flag(a.hidden) || col.width === 0) col.hidden = true
    if (a.style !== undefined) Object.assign(col, look(read.xfs[Number(a.style)], read.issues, where))
    if (Number(a.outlineLevel ?? 0) > 0) read.issues.add('outline', 'grouped rows and columns are not in the model', where)
    if (Object.keys(col).length) lines.push([lo, hi, col])
  }
  return spans(lines, 'c')
}

function merges(rest: string, sheet: string, issues: Issues): string[] | undefined {
  const xs: Rect[] = []
  for (const e of elements(element(rest, 'mergeCells')?.body ?? '', 'mergeCell')) {
    const ref = attributes(e.attrs).ref?.toUpperCase()
    const x = ref ? parseMerge(ref) : null
    if (!x) continue
    xs.push(x)
    const o = overlapping(xs)
    if (o) {
      xs.pop()
      issues.add('part', 'a merged area overlapping another is left out', { sheet, cell: ref })
    }
  }
  return xs.length ? xs.map(mergeText) : undefined
}

/** Whether an autofilter, a sheet's or a table's, has criteria: a column with a filter of any kind. */
export function hasCriteria(filter: Element | undefined): boolean {
  return !!filter?.body && elements(filter.body, 'filterColumn').some((c) => !!c.body && /<[\w.:-]+[\s/>]/.test(c.body))
}

/**
 * Excel's filter mode: an autofilter with criteria, the sheet's or a table's, or the mark of an advanced filter (no
 * autofilter). The mark beside an autofilter without criteria is not taken, as Excel does not take it. The range and
 * the criteria are reported.
 */
function filterMode(head: string, tail: string, src: SheetSource, issues: Issues): boolean {
  const where = { sheet: src.name }
  const own = element(tail, 'autoFilter')
  const mark = flag(attributes(element(head, 'sheetPr')?.attrs ?? '').filterMode)
  const tables = src.tables.some((x) => hasCriteria(element(x, 'autoFilter')))
  const criteria = hasCriteria(own)
  if (own) issues.add('autoFilter', criteria ? "an autofilter's range, criteria and sort are not in the model: the sheet is read as filtered, the rows it hid as hidden" : 'an autofilter without criteria (its buttons and sort) is not in the model', where)
  else if (mark && !tables) issues.add('autoFilter', 'an advanced filter is not in the model: the sheet is read as filtered, the rows it hid as hidden', where)
  if (tables) issues.add('autoFilter', "a table's filter criteria are not in the model: the sheet is read as filtered, the rows they hid as hidden", where)
  return criteria || tables || (mark && !own)
}

function features(head: string, tail: string, src: SheetSource, issues: Issues) {
  const outside = head + tail
  const where = { sheet: src.name }
  for (const [local, code, message] of FEATURES) if (has(outside, local)) issues.add(code, message, where)
  // the sheet's default column width: 8 digits and the padding, as the model assumes, or another
  const fp = attributes(element(head, 'sheetFormatPr')?.attrs ?? '')
  const width = fp.defaultColWidth !== undefined ? Number(fp.defaultColWidth) : Number(fp.baseColWidth ?? 8) + 0.84
  if (!(width >= 8 && width < 9.5)) issues.add('defaultWidth', `the sheet's default column width (${Math.round(width * 100) / 100} in xlsx units) is not in the model, which takes Excel's standard of 8 digits`, where)
  const setup = element(tail, 'pageSetup')
  const sa = setup ? attributes(setup.attrs) : {}
  if (sa.orientation === 'landscape' || sa.scale !== undefined || sa.fitToWidth !== undefined || sa.fitToHeight !== undefined) issues.add('print', 'print settings (headers and footers, page breaks, page setup) are not in the model', where)
  for (const r of src.rels) if (LINKED[r.kind]) issues.add(...LINKED[r.kind], where)
}

export function readSheet(src: SheetSource, ctx: CellsContext): SheetRead {
  const { xml, name } = src
  const open = xml.search(/<(?:[\w.-]+:)?sheetData[\s>/]/)
  const closeAt = open < 0 ? -1 : xml.indexOf('sheetData>', open + 10)
  const head = open < 0 ? xml : xml.slice(0, open)
  const dataEnd = closeAt < 0 ? open : closeAt + 10
  const data = open < 0 ? '' : xml.slice(open, dataEnd)
  const tail = open < 0 ? '' : xml.slice(dataEnd)
  const read = readCells(data, ctx)
  for (const s of read.shared.values()) if (s.master !== undefined) read.cells[s.master] = { ...read.cells[s.master], formula: s.text }
  const converted = convert(src, read, ctx.issues)
  // copies are pasted from the converted master and take what it reports
  spreadShared(src, read, ctx.issues)
  for (const s of read.shared.values()) {
    const master = s.master === undefined ? undefined : converted.get(s.master)
    if (master)
      for (const cell of s.copies) {
        if (read.cells[cell]?.formula === undefined) {
          ctx.issues.add('formula', 'a copy of a shared formula that a paste could not make is left out', { sheet: name, cell })
          continue
        }
        reportFormula(master, ctx.issues, { sheet: name, cell })
      }
  }
  arrays(read)
  for (const cell of sortAddresses(read.errors)) ctx.issues.add('errorConstant', 'an error typed into a cell is read as a formula of that error', { sheet: name, cell })
  features(head, tail, src, ctx.issues)
  const sheet: SheetInput = { name, cells: read.cells }
  if (filterMode(head, tail, src, ctx.issues)) sheet.filtered = true
  const cols = columns(head, ctx)
  if (cols) sheet.columns = cols
  const rows = spans<RowInput>([...read.rows].map(([r, v]) => [r - 1, r - 1, v]), 'r')
  if (rows) sheet.rows = rows
  const m = merges(tail, name, ctx.issues)
  if (m) sheet.merges = m
  const tables = readTables(src.tables, name, read.cells, m, ctx.issues)
  if (tables.length) sheet.tables = tables
  readCellParts({ sheet: name, head, tail, rels: src.rels, ...src.parts }, read.cells, ctx.issues)
  const frozen = readFrozen(head, name, ctx.issues)
  if (frozen) sheet.frozen = frozen
  return { sheet, cached: read.cached }
}
