/**
 * A worksheet part as a sheet of the book: cells (see `cells.ts`), formulas in the engine's spelling (those of the
 * old kind as the dynamic formulas Excel 365 shows for them, `@` where Excel puts it), shared formulas spread to their
 * cells as a copy and paste spreads them (`editStructure`'s copyRange), array formulas of the old kind over their area,
 * the cells an array spilled into left to the stored values, columns and rows with their outline levels, merged areas,
 * the filter mode, notes, hyperlinks and frozen panes (`parts.ts`), data validation (`validation.ts`), outline settings,
 * and what the sheet has that the model does not, reported.
 */
import type { CellInput, CellValue, ColumnInput, OutlineSettings, RowInput, SheetInput } from '../../contract.js'
import { cellEdits, edited, LookTable, NONE, type Layer } from '../../looks/layer.js'
import { listOf, lookAlone } from '../../looks/sheet.js'
import { overlapping, parseMerge } from '../../layout/merges.js'
import { lines, spanKey } from '../../layout/spans.js'
import { formatAddress, MAX_COLS, parseAddress, readAddress, sortAddresses } from '../../model.js'
import { cellCount, cellsWithin, cornersText, parseRange, type Rect } from '../../structure/area.js'
import { copyRange } from '../../structure/copy.js'
import type { Ast } from '../../syntax/ast.js'
import { canonical } from '../../syntax/print.js'
import type { Issues, XlsxMessage } from '../issues.js'
import type { XlsxCachedCell, XlsxIssueCode } from '../types.js'
import { attributes, element, elements, flag, has, ownEntry, type Element } from '../xml/scan.js'
import { look, outlineOf, readCells, type CellsContext, type SheetCells } from './cells.js'
import { readFormula, reportFormula, type FormulaOptions, type ReadFormula } from './formulas.js'
import type { Rel } from './package.js'
import { readCellParts, readFrozen, type PartsSource } from './parts.js'
import { readTables } from './tables.js'
import { outsideConditional, readConditional } from './conditional.js'
import type { Dxf } from '../styles/dxf.js'
import type { Palette } from '../styles/color.js'
import { readValidations } from './validation.js'

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
  /** The book's differential formats and colours, which conditional formats use. */
  cond: { dxfs: Dxf[]; palette: Palette }
  /** How its formulas are read: the book's links to other workbooks they may refer to. */
  formulas: FormulaOptions
}

export interface SheetRead {
  sheet: SheetInput
  cached: Record<string, XlsxCachedCell>
}

/** Elements of a worksheet the model has no place for, outside `<sheetData>`. */
const FEATURES: [local: string, code: XlsxIssueCode, message: XlsxMessage][] = [
  ['sheetProtection', 'protection', { key: 'sheetProtection' }],
  ['protectedRanges', 'protection', { key: 'sheetProtection' }],
  ['drawing', 'drawing', { key: 'drawings' }],
  ['picture', 'drawing', { key: 'drawings' }],
  ['oleObjects', 'drawing', { key: 'drawings' }],
  ['controls', 'drawing', { key: 'drawings' }],
  ['sparklineGroups', 'sparkline', { key: 'sparklines' }],
  ['headerFooter', 'print', { key: 'printSettings' }],
  ['rowBreaks', 'print', { key: 'printSettings' }],
  ['colBreaks', 'print', { key: 'printSettings' }],
]

/** Parts a sheet links to that the model has no place for, by relationship kind. */
const LINKED: Record<string, [XlsxIssueCode, XlsxMessage]> = {
  pivotTable: ['pivotTable', { key: 'pivotTables' }],
}

/** Shared formulas in each of their cells: the master's formula copied and pasted over them. */
function spreadShared(src: SheetSource, read: SheetCells, issues: Issues) {
  for (const s of read.shared.values()) {
    if (s.master === undefined) {
      for (const c of s.copies) issues.add('formula', { key: 'sharedWithoutMaster' }, { sheet: src.name, cell: c })
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
    const targets = cellCount(box) > 4 * at.length ? s.copies : [cornersText(box)]
    const book = { sheets: src.sheets.map((name) => ({ name, cells: name === src.name ? { [s.master!]: { formula } } : {} })) }
    for (const range of targets) {
      // the pure copy and paste, without editStructure's checks and values (a paste from a filtered sheet computes the book)
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
    const f = readFormula(cell.formula, issues, { sheet: src.name, cell: addr }, src.formulas)
    const dynamic = f.ast && read.old.has(addr) ? src.legacy(f.ast, src.name) : undefined
    cell.formula = dynamic ? canonical(dynamic) : f.formula
    out.set(addr, f)
  }
  return out
}

/**
 * The cells an array formula spilled into hold stored values: they go to `cached` with the formula's address, keeping
 * their look; the formula's value gets the range, as a result has them (an old kind's area always, a dynamic array's
 * when it spilled).
 */
function arrays(read: SheetCells): string[] {
  const keys = () => Object.keys(read.cells)
  // the cells left with their look alone
  const looked: string[] = []
  for (const a of read.arrays) {
    const r = parseRange(a.ref)
    if (!r) continue
    const anchor = read.cells[a.anchor]
    if (!a.dynamic && anchor?.formula !== undefined) anchor.array = a.ref
    const top = read.cached[a.anchor]
    if (top && (!a.dynamic || r.r2 > r.r1 || r.c2 > r.c1)) top.spill = cornersText(r)
    for (const addr of cellsWithin(r, keys)) {
      const cell = read.cells[addr]
      // an error the array spilled is stored as the cell's value, which the reader took for a formula of it
      const spilled = read.errors.delete(addr)
      if (addr === a.anchor || !cell || (cell.formula !== undefined && !spilled)) continue
      const stored = cell.value !== undefined ? { value: cell.value as CellValue } : read.cached[addr]
      if (stored) read.cached[addr] = { value: stored.value, spillOf: a.anchor }
      const { value: _, formula: __, ...rest } = cell
      if (Object.keys(rest).length) {
        read.cells[addr] = rest
        looked.push(addr)
      } else delete read.cells[addr]
    }
  }
  return looked
}

/** Where a sheet's summary rows and columns are (`<outlinePr>`), as the file says; automatic outline styles are reported. */
function outline(head: string, sheet: string, issues: Issues): OutlineSettings | undefined {
  const pr = element(element(head, 'sheetPr')?.body ?? '', 'outlinePr')
  if (!pr) return undefined
  const a = attributes(pr.attrs)
  if (flag(a.applyStyles)) issues.add('outline', { key: 'outlineStyles' }, { sheet })
  const out: OutlineSettings = {}
  if (a.summaryBelow !== undefined) out.summaryBelow = flag(a.summaryBelow)
  if (a.summaryRight !== undefined) out.summaryRight = flag(a.summaryRight)
  return Object.keys(out).length ? out : undefined
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
    // a width Excel fitted, which typing widens again
    if (col.width !== undefined && flag(a.bestFit)) col.bestFit = true
    // Excel hides a column of width 0 whatever its flag
    if (flag(a.hidden) || col.width === 0) col.hidden = true
    if (a.style !== undefined) Object.assign(col, look(read.xfs[Number(a.style)], read.issues, where))
    Object.assign(col, outlineOf(a))
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
      issues.add('part', { key: 'mergeOverlap' }, { sheet, cell: ref })
    }
  }
  return xs.length ? xs.map(cornersText) : undefined
}

/** Whether an autofilter, a sheet's or a table's, has criteria: a column with a filter of any kind. */
function hasCriteria(filter: Element | undefined): boolean {
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
  if (own) issues.add('autoFilter', { key: criteria ? 'autoFilterCriteria' : 'autoFilterPlain' }, where)
  else if (mark && !tables) issues.add('autoFilter', { key: 'advancedFilter' }, where)
  if (tables) issues.add('autoFilter', { key: 'tableFilter' }, where)
  return criteria || tables || (mark && !own)
}

function features(head: string, tail: string, src: SheetSource, issues: Issues) {
  const outside = head + tail
  const where = { sheet: src.name }
  for (const [local, code, message] of FEATURES) if (has(outside, local)) issues.add(code, message, where)
  // the sheet's default column width: 8 digits and the padding, as the model assumes, or another
  const fp = attributes(element(head, 'sheetFormatPr')?.attrs ?? '')
  const width = fp.defaultColWidth !== undefined ? Number(fp.defaultColWidth) : Number(fp.baseColWidth ?? 8) + 0.84
  if (!(width >= 8 && width < 9.5)) issues.add('defaultWidth', { key: 'defaultWidth', width: Math.round(width * 100) / 100 }, where)
  const setup = element(tail, 'pageSetup')
  const sa = setup ? attributes(setup.attrs) : {}
  if (sa.orientation === 'landscape' || sa.scale !== undefined || sa.fitToWidth !== undefined || sa.fitToHeight !== undefined) issues.add('print', { key: 'printSettings' }, where)
  for (const r of src.rels) {
    const linked = ownEntry(LINKED, r.kind)
    if (linked) issues.add(...linked, where)
  }
}

/** The addresses of the cells of a look of a layer in a rectangle. */
function* addressesOf(layer: Layer, x: Rect): Generator<string> {
  for (const p of layer.pieces(x)) for (let r = p.r1; r <= p.r2; r++) for (let c = p.c1; c <= p.c2; c++) yield formatAddress(r, c)
}

/**
 * The sheet's looks with its cells as the reader left them: a cell left with a look alone (one an array spilled into, a
 * shared formula's copy without its formula) goes to the looks, and a cell given content over a look (a shared formula's
 * copy, a note, a link, a table's header) takes the look as its own, as the cell's entry of a look would have, and its
 * place among the cells the file stores (`stored`), as the entry had. Without looks, only the cells that may be left with
 * a look alone (`maybe`) are looked at. The cells and the looks.
 */
function withEntries(layer: Layer, cells: Record<string, CellInput>, stored: Set<string> | null, maybe: Iterable<string>): [Record<string, CellInput>, Layer] {
  const looks = new LookTable(layer.looks)
  const edits: [number, number, number][] = []
  const p = { r: 0, c: 0 }
  const made: [number, string][] = []
  for (const a of layer.empty ? maybe : Object.keys(cells)) {
    const cell = cells[a]
    if (!cell) continue
    const alone = lookAlone(cell)
    if (!alone && layer.empty) continue
    readAddress(a, p)
    if (alone) {
      edits.push([p.r, p.c, looks.of(cell)])
      delete cells[a]
      continue
    }
    const look = layer.lookAt(p.r, p.c)
    if (!look) continue
    cells[a] = { ...look, ...cell }
    edits.push([p.r, p.c, NONE])
    if (!stored?.has(a)) made.push([p.r * MAX_COLS + p.c, a])
  }
  return [made.length ? inFileOrder(cells, stored!, made) : cells, edited(layer, cellEdits(edits), looks)]
}

/** The cells with those `made` (by place) among the ones the file stores, by place, before the others the reader made. */
function inFileOrder(cells: Record<string, CellInput>, stored: Set<string>, made: [number, string][]): Record<string, CellInput> {
  made.sort((a, b) => a[0] - b[0])
  const gone = new Set(made.map(([, a]) => a))
  const out: Record<string, CellInput> = {}
  const p = { r: 0, c: 0 }
  let i = 0
  const keys = Object.keys(cells).filter((a) => !gone.has(a))
  for (const a of keys.filter((x) => stored.has(x))) {
    readAddress(a, p)
    for (; i < made.length && made[i][0] < p.r * MAX_COLS + p.c; i++) out[made[i][1]] = cells[made[i][1]]
    out[a] = cells[a]
  }
  for (; i < made.length; i++) out[made[i][1]] = cells[made[i][1]]
  for (const a of keys) if (!stored.has(a)) out[a] = cells[a]
  return out
}

export function readSheet(src: SheetSource, ctx: CellsContext): SheetRead {
  const { xml, name } = src
  const open = xml.search(/<(?:[\w.-]+:)?sheetData[\s>/]/)
  const closeAt = open < 0 ? -1 : xml.indexOf('sheetData>', open + 10)
  const head = open < 0 ? xml : xml.slice(0, open)
  const dataEnd = closeAt < 0 ? open : closeAt + 10
  const data = open < 0 ? '' : xml.slice(open, dataEnd)
  const tail = open < 0 ? '' : xml.slice(dataEnd)
  const cols = columns(head, ctx)
  const read = readCells(data, ctx, lines(cols, 'c'))
  // the looks, and with them the cells the file stores, among which a cell given content over a look takes its place
  const layer = read.looks.build()
  const stored = layer.empty ? null : new Set(Object.keys(read.cells))
  for (const s of read.shared.values()) if (s.master !== undefined) read.cells[s.master] = { ...read.cells[s.master], formula: s.text }
  const converted = convert(src, read, ctx.issues)
  // copies are pasted from the converted master and take what it reports
  spreadShared(src, read, ctx.issues)
  for (const s of read.shared.values()) {
    const master = s.master === undefined ? undefined : converted.get(s.master)
    if (master)
      for (const cell of s.copies) {
        if (read.cells[cell]?.formula === undefined) {
          ctx.issues.add('formula', { key: 'sharedCopy' }, { sheet: name, cell })
          continue
        }
        reportFormula(master, ctx.issues, { sheet: name, cell })
      }
  }
  const looked = arrays(read)
  for (const cell of sortAddresses(read.errors)) ctx.issues.add('errorConstant', { key: 'errorTyped' }, { sheet: name, cell })
  features(head, tail, src, ctx.issues)
  const sheet: SheetInput = { name, cells: read.cells }
  if (filterMode(head, tail, src, ctx.issues)) sheet.filtered = true
  if (cols) sheet.columns = cols
  const rows = spans<RowInput>([...read.rows].map(([r, v]) => [r - 1, r - 1, v]), 'r')
  if (rows) sheet.rows = rows
  const m = merges(tail, name, ctx.issues)
  if (m) sheet.merges = m
  const tables = readTables(src.tables, name, read.cells, m, ctx.issues, src.formulas)
  if (tables.length) sheet.tables = tables
  readCellParts({ sheet: name, head, tail, rels: src.rels, ...src.parts, looked: (x) => addressesOf(layer, x) }, read.cells, ctx.issues)
  const [cells, looks] = withEntries(layer, read.cells, stored, [...looked, ...[...read.shared.values()].flatMap((s) => s.copies)])
  sheet.cells = cells
  if (!looks.empty) sheet.looks = listOf(looks)
  const frozen = readFrozen(head, name, ctx.issues)
  if (frozen) sheet.frozen = frozen
  const cf = readConditional(tail, name, src.cond.dxfs, src.cond.palette, ctx.issues, src.formulas)
  if (cf) sheet.conditionalFormats = cf
  const validations = readValidations(outsideConditional(tail), name, src.sheets, ctx.issues, src.formulas)
  if (validations) sheet.validations = validations
  const settings = outline(head, name, ctx.issues)
  if (settings) sheet.outline = settings
  return { sheet, cached: read.cached }
}
