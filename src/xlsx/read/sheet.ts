/**
 * A worksheet part as a sheet of the book: cells (see `cells.ts`), formulas in the engine's spelling (those of the
 * old kind as the dynamic formulas Excel 365 shows for them, `@` where Excel puts it), shared formulas spread to their
 * cells as a copy and paste spreads them (`restructure`'s copyRange), array formulas of the old kind over their area,
 * the cells an array spilled into left to the stored values, columns, rows, merged areas, and what the sheet has that
 * the model does not, reported.
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
import { attributes, element, elements, flag, has } from '../xml/scan.js'
import { look, readCells, type CellsContext, type SheetCells } from './cells.js'
import { readFormula, reportFormula, type ReadFormula } from './formulas.js'
import type { Rel } from './package.js'
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
  ['hyperlinks', 'hyperlink', 'hyperlinks are not in the model'],
  ['autoFilter', 'autoFilter', 'an autofilter and its sort are not in the model'],
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
  comments: ['comment', 'notes and comments are not in the model'],
  threadedComment: ['comment', 'notes and comments are not in the model'],
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
    const box = { r1: Math.min(...at.map((a) => a.row)), r2: Math.max(...at.map((a) => a.row)), c1: Math.min(...at.map((a) => a.col)), c2: Math.max(...at.map((a) => a.col)) }
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
  for (const [addr, cell] of Object.entries(read.cells)) {
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
    if (flag(a.hidden)) {
      col.width = 0
      read.issues.add('hidden', 'hidden rows and columns are read with a size of 0', where)
    } else if (a.width !== undefined && flag(a.customWidth)) col.width = Math.min(255, Number(a.width))
    if (a.style !== undefined) Object.assign(col, look(read.xfs[Number(a.style)]))
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

function features(head: string, tail: string, src: SheetSource, issues: Issues) {
  const outside = head + tail
  const where = { sheet: src.name }
  for (const [local, code, message] of FEATURES) if (has(outside, local)) issues.add(code, message, where)
  // the sheet's default column width: 8 digits and the padding, as the model assumes, or another
  const fp = attributes(element(head, 'sheetFormatPr')?.attrs ?? '')
  const width = fp.defaultColWidth !== undefined ? Number(fp.defaultColWidth) : Number(fp.baseColWidth ?? 8) + 0.84
  if (!(width >= 8 && width < 9.5)) issues.add('defaultWidth', `the sheet's default column width (${Math.round(width * 100) / 100} in xlsx units) is not in the model, which takes Excel's standard of 8 digits`, where)
  const pane = element(head, 'pane')
  const state = pane && attributes(pane.attrs).state
  if (state === 'frozen' || state === 'frozenSplit' || state === 'split') issues.add('frozenPanes', 'frozen and split panes are not in the model', where)
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
  const cols = columns(head, ctx)
  if (cols) sheet.columns = cols
  const rows = spans<RowInput>([...read.rows].map(([r, v]) => [r - 1, r - 1, v]), 'r')
  if (rows) sheet.rows = rows
  const m = merges(tail, name, ctx.issues)
  if (m) sheet.merges = m
  const tables = readTables(src.tables, name, read.cells, m, ctx.issues)
  if (tables.length) sheet.tables = tables
  return { sheet, cached: read.cached }
}
