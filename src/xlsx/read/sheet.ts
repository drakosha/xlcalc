/**
 * A worksheet part as a sheet of the book: cells (see `cells.ts`), formulas in the engine's spelling, shared
 * formulas spread to their cells as a copy and paste spreads them (`restructure`'s copyRange), the cells an
 * array spilled into left to the stored values, columns, rows, merged areas, and what the sheet has that the
 * model does not, reported.
 */
import type { CellValue, ColumnInput, RowInput, SheetInput } from '../../contract.js'
import { mergeText, overlapping, parseMerge } from '../../layout/merges.js'
import { spanKey } from '../../layout/spans.js'
import { formatAddress, parseAddress } from '../../model.js'
import { parseRange, type Rect } from '../../structure/area.js'
import { restructure } from '../../structure/index.js'
import type { Ast } from '../../syntax/ast.js'
import type { Issues } from '../issues.js'
import type { XlsxIssueCode } from '../types.js'
import { attributes, element, elements, flag, has } from '../xml/scan.js'
import { look, readCells, type CellsContext, type SheetCells } from './cells.js'
import { readFormula } from './formulas.js'
import type { Rel } from './package.js'

export interface SheetSource {
  name: string
  xml: string
  rels: Rel[]
  /** The names of the book's sheets: shared formulas may reference any of them. */
  sheets: string[]
}

export interface SheetRead {
  sheet: SheetInput
  cached: Record<string, CellValue>
  /** Syntax trees of the formulas of the old kind, by cell (`intersect.ts` checks them). */
  old: Map<string, Ast>
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
  ['tableParts', 'table', 'tables (list objects) are not in the model: their cells are read as plain cells'],
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
      const r = restructure(book, { op: 'copyRange', from: { sheet: src.name, range: s.master }, to: { sheet: src.name, range } })
      const pasted = r.ok ? r.book.sheets.find((x) => x.name === src.name)!.cells : {}
      for (const c of s.copies) {
        const f = pasted[c]?.formula
        if (f !== undefined) read.cells[c] = { ...read.cells[c], formula: f }
      }
    }
  }
}

/** The formulas of `addrs` in the engine's spelling; the trees of the old kind kept for the intersection check. */
function convert(src: SheetSource, read: SheetCells, issues: Issues, old: Map<string, Ast>, addrs: Iterable<string>) {
  for (const addr of addrs) {
    const cell = read.cells[addr]
    if (cell?.formula === undefined) continue
    const f = readFormula(cell.formula, issues, { sheet: src.name, cell: addr })
    cell.formula = f.formula
    if (f.ast && read.old.has(addr)) old.set(addr, f.ast)
  }
}

/** The cells an array formula spilled into hold stored values: they go to `cached`, keeping their look. */
function arrays(src: SheetSource, read: SheetCells, issues: Issues) {
  const keys = () => Object.keys(read.cells)
  for (const a of read.arrays) {
    if (!a.dynamic) issues.add('legacyArray', 'an array formula of the old kind is read as a dynamic array formula: its array spills from the top-left cell', { sheet: src.name, cell: a.ref })
    const r = parseRange(a.ref)
    if (!r) continue
    for (const addr of cellsOf(r, keys)) {
      const cell = read.cells[addr]
      if (addr === a.anchor || !cell || cell.formula !== undefined) continue
      if (cell.value !== undefined) read.cached[addr] = cell.value as CellValue
      const { value: _, ...rest } = cell
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
  const old = new Map<string, Ast>()
  for (const s of read.shared.values()) if (s.master !== undefined) read.cells[s.master] = { ...read.cells[s.master], formula: s.text }
  convert(src, read, ctx.issues, old, Object.keys(read.cells))
  // copies are pasted from the converted master, then checked as their own formulas
  spreadShared(src, read, ctx.issues)
  convert(src, read, ctx.issues, old, [...read.shared.values()].flatMap((s) => (s.master === undefined ? [] : s.copies)))
  arrays(src, read, ctx.issues)
  features(head, tail, src, ctx.issues)
  const sheet: SheetInput = { name, cells: read.cells }
  const cols = columns(head, ctx)
  if (cols) sheet.columns = cols
  const rows = spans<RowInput>([...read.rows].map(([r, v]) => [r - 1, r - 1, v]), 'r')
  if (rows) sheet.rows = rows
  const m = merges(tail, name, ctx.issues)
  if (m) sheet.merges = m
  return { sheet, cached: read.cached, old }
}
