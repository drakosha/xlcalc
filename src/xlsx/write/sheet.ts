/**
 * A worksheet part from a sheet of the book. Formulas are written as dynamic array formulas (`cm` with the XLDAPR
 * metadata, `t="array"` over the range their array spilled into), as Excel 365 saves them: the engine computes
 * every formula by the rules of dynamic arrays, and so does Excel then. An array formula of the old kind
 * (`CellInput.array`) is written over its area without the metadata, as Excel writes one entered with
 * Ctrl+Shift+Enter, and a formula in a table's totals row as a plain one, the only kind Excel takes there. Values
 * come from the result: a formula's value, and the cells its array spilled into; a formula a circular reference stops
 * stores its `held` where the result gives no value for it, as Excel opens it again holding that. A formula the file cannot hold (it
 * does not parse, or names a sheet, a table, a column or a link to another workbook the book does not have) is written
 * as its value. A filtered sheet gets criteria that every row meets (`KEEP_ALL`): on its first table with a header row,
 * else on an autofilter of its own over its cells, with Excel's filter mark. Every cell is written with the whole look it
 * shows (its own style and format over its row's or column's, `layout/lines.ts`), as Excel shows a stored cell by its own
 * cell format alone, and so are the empty cells where a row crosses a column that Excel would otherwise show by the row's
 * whole format. Notes, hyperlinks and frozen panes are written by `parts.ts`, data validation by `validation.ts`; rows
 * and columns carry their outline levels.
 */
import type { CellInput, CellLink, CellNote, CellValue, ColumnInput, ErrorCode, ExternalBook, RowInput, SheetInput } from '../../contract.js'
import { sheetViewXml } from '../view.js'
import { lookOf, sheetLines, type Look } from '../../layout/lines.js'
import { normalized } from '../../layout/outline.js'
import { lines, type Line } from '../../layout/spans.js'
import { dateToSerial, formatAddress, indexToCol, MAX_COLS, parseAddress, readAddress } from '../../model.js'
import { NONE, type Layer } from '../../looks/layer.js'
import { inputIn, layerOf } from '../../looks/sheet.js'
import { meets, parseRange, rectCells } from '../../structure/area.js'
import { ERROR_TEXT, FormulaError, type Ast } from '../../syntax/ast.js'
import { parse } from '../../syntax/parse.js'
import { sheetKey, walk } from '../../syntax/print.js'
import { bindTables } from '../../syntax/check.js'
import { bindLinks } from '../../syntax/links.js'
import { toFile } from '../formulas/to-file.js'
import type { Issues } from '../issues.js'
import { escapeAttr, escapeText } from '../xml/text.js'
import { commentsXml, hyperlinksXml, paneXml, vmlXml } from './parts.js'
import { conditionalXml } from './conditional.js'
import { validationsXml } from './validation.js'
import type { StyleWriter } from './styles.js'
import type { ResultCell } from './snapshot.js'
import { clip, HEAD, MAIN, MAX_TEXT } from './xml.js'

/**
 * The default width in xlsx units (8 digits and the padding, as Excel for Mac writes it): a column that has a look
 * but no width of its own needs one, as Excel reads a column without a width as hidden.
 */
const DEFAULT_WIDTH = '8.83203125'

/**
 * The criteria of the autofilter a filtered sheet gets: a cell equal to nothing or not equal to it, which every row
 * meets. Excel takes the sheet as filtered by them and leaves its rows as the file has them; its Reapply shows every row.
 */
export const KEEP_ALL = '<filterColumn colId="0"><customFilters><customFilter val=""/><customFilter operator="notEqual" val=""/></customFilters></filterColumn>'

/** What a formula of the book may name in its file. */
export interface FileNames {
  /** Upper-case names (`sheetKey`) of the book's sheets. */
  sheets: Set<string>
  /** The book's tables with their columns. */
  tables: { name: string; columns: string[] }[]
  /** The book's links to other workbooks, which `[1]Sheet1!A1` names. */
  externalBooks?: ExternalBook[]
}

export interface SheetContext extends FileNames {
  styles: StyleWriter
  /** The index of a string in the shared strings table. */
  string(s: string): number
  issues: Issues
  mode: 'formulas' | 'values'
  /** The sheet's cells of the result, when the file stores values of formulas. */
  result?: Record<string, ResultCell>
  /** The value metadata index (`vm`, 1-based) of an error Excel stores as a rich value: #SPILL!, #CALC!. */
  richError(code: 'SPILL' | 'CALC'): number
  /** The book's table holding a cell of this sheet. */
  tableAt(addr: string): string | undefined
  /** A cell of a table's totals row, where Excel takes only a formula of the old kind (no array formula). */
  totalsAt(addr: string): boolean
  /** The relationship ids of the sheet's table parts. */
  tableParts: string[]
  /** The day 0 of the book's date system in the 1900 system (`epochOf`): a date's serial counts from it. */
  epoch: number
  /** A filtered sheet's table whose autofilter holds the criteria (`KEEP_ALL`), in place of an autofilter of the sheet's own, as Excel keeps a table's filter. */
  filterTable?: string
  /** The first block of the VML shape ids of the sheet's notes (`vmlXml`). */
  noteBlock: number
  /** The sheet is the book's active one (`workbookView activeTab`), its tab selected unless the saved view specifies it. */
  active: boolean
}

/** A relationship of the sheet after its tables': its id is `rId` and its number among all the sheet's relationships. */
export interface SheetRel {
  type: 'hyperlink' | 'comments' | 'vmlDrawing'
  target: string
}

export interface SheetWritten {
  xml: string
  /** Whether a formula was written (the book then needs the dynamic array metadata). */
  formulas: boolean
  /** The sheet's relationships after its tables' (hyperlinks' addresses are external). */
  rels: SheetRel[]
  /** The comments part and VML drawing of the sheet's notes, and their count. */
  notes?: { comments: string; vml: string; count: number }
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
    const s = clip(v, MAX_TEXT, ctx.issues, { key: 'textCut' }, where)
    return shared ? { t: ' t="s"', v: String(ctx.string(s)) } : { t: ' t="str"', v: escapeText(s) }
  }
  return errorXml(v.error, () => ctx.issues.add('errorValue', { key: 'cycleZero' }, where), ctx.richError)
}

/**
 * An error as a file's cell stores it, its attributes and the content of `<v>`: #SPILL! and #CALC! as `#VALUE!`, with a
 * rich value (`vm`) that names them where `richError` gives one, as Excel 365 stores them in a sheet; the engine's
 * #CYCLE! as 0, as Excel stores a circular reference, which `cycle` reports.
 */
export function errorXml(code: ErrorCode, cycle: () => void, richError?: SheetContext['richError']): { t: string; v: string } {
  if (code === 'CYCLE') {
    cycle()
    return { t: '', v: '0' }
  }
  if (code === 'SPILL' || code === 'CALC') return { t: ` t="e"${richError ? ` vm="${richError(code)}"` : ''}`, v: '#VALUE!' }
  return { t: ' t="e"', v: code === 'ERROR' ? '#NULL!' : ERROR_TEXT[code] }
}

/**
 * The formula's tree when the file can hold it: it parses, names only the book's sheets, its structured references name
 * the book's tables and columns (a file names a reference's table always: `[@Qty]` is written `Sales[[#This Row],[Qty]]`),
 * and its references to other workbooks the book's links and their sheets.
 */
export function fileTree(formula: string, names: FileNames, enclosing?: string): Ast | null {
  let ast: Ast
  try {
    ast = parse(formula)
  } catch (e) {
    if (e instanceof FormulaError) return null
    throw e
  }
  const bound = bindTables(ast, formula, { tables: names.tables, ...(enclosing ? { table: enclosing } : {}) })
  if (!bound.ok) return null
  ast = bound.ast
  let ok = true
  let linked = false
  const known = (sheet: string | null | undefined) => sheet == null || names.sheets.has(sheetKey(sheet))
  walk(ast, (n) => {
    // a reference into another workbook, one Excel lost too (`[1]Sheet1!#REF!`, of any sheet of its link)
    if (n.k === 'xref' || n.k === 'xname' || (n.k === 'err' && n.book !== undefined)) linked = true
    else if (n.k === 'ref3d' ? !known(n.first) || !known(n.last) : (n.k === 'ref' || n.k === 'err' || n.k === 'name') && !known(n.sheet)) ok = false
  })
  if (!ok || !linked) return ok ? ast : null
  const links = bindLinks(ast, names.externalBooks, [...names.sheets])
  return links.ok ? links.ast : null
}

export function writeSheet(sheet: SheetInput, index: number, ctx: SheetContext): SheetWritten {
  const { styles, issues, mode, result } = ctx
  // each row's cells' XML by column and its columns written, and the bounds of them all: no map of every cell's place
  const xmlAt: string[][] = [], colsOf: number[][] = []
  const extent = { r1: Infinity, r2: -1, c1: MAX_COLS, c2: 0, n: 0 }
  const rowsUsed = new Set<number>()
  let formulas = false
  const put = (addr: string, xml: string) => {
    const { row, col } = parseAddress(addr)
    const xs = (xmlAt[row] ??= [])
    if (xs[col] === undefined) {
      ;(colsOf[row] ??= []).push(col)
      extent.n++
      extent.r1 = Math.min(extent.r1, row)
      extent.r2 = Math.max(extent.r2, row)
      extent.c1 = Math.min(extent.c1, col)
      extent.c2 = Math.max(extent.c2, col)
    }
    xs[col] = xml
    rowsUsed.add(row)
  }
  const where = (cell: string) => ({ sheet: sheet.name, cell })
  const sheetLook = sheetLines(sheet)
  // a sheet whose rows and columns have no look writes each cell's own
  const lined = [sheet.rows, sheet.columns].some((m) => m && Object.values(m).some((e) => e.style !== undefined || e.numberFormat !== undefined))
  const sAttr = (addr: string, c: CellInput | undefined) => {
    let look: Look | undefined = c
    if (lined) {
      const { row, col } = parseAddress(addr)
      look = sheetLook.look(row, col, c)
    }
    const xf = look ? styles.xf(look.style, look.numberFormat) : 0
    return xf ? ` s="${xf}"` : ''
  }
  // `cm` marks a dynamic array formula
  // the cells of the sheet's looks numbered among its entries, as entries of theirs would be (`numberLooks`)
  const layer = layerOf(sheet)
  const layerXf = layer ? numberLooks(sheet, layer, sAttr) : null
  const cellXml = (addr: string, c: CellInput | undefined, v: CellValue | undefined, f = '', dynamic = f !== '') => {
    const cm = dynamic ? ' cm="1"' : ''
    if (v === undefined) return `<c r="${addr}"${sAttr(addr, c)}${cm}${f ? `>${f}</c>` : '/>'}`
    const x = valueXml(v, ctx, where(addr), !f)
    return `<c r="${addr}"${sAttr(addr, c)}${x.t}${cm}>${f}<v>${x.v}</v></c>`
  }
  const notes = new Map<string, CellNote>()
  const links = new Map<string, CellLink>()
  for (const addr of Object.keys(sheet.cells)) {
    const c = sheet.cells[addr]
    if (c.meta !== undefined) issues.add('meta', { key: 'cellMeta' }, where(addr))
    if (c.note) notes.set(addr, c.note)
    if (c.link) links.set(addr, c.link)
    const res = result?.[addr]
    if (c.formula === undefined) {
      const v = c.value === undefined || c.value === null ? undefined : typeof c.value === 'object' ? dateToSerial(c.value.date) - ctx.epoch : c.value
      put(addr, cellXml(addr, c, v))
      continue
    }
    // a formula a circular reference stops stores the value it holds where the result gives none, or the engine's #CYCLE!
    const value = c.held !== undefined && (res === undefined || (typeof res.value === 'object' && res.value?.error === 'CYCLE')) ? c.held : res?.value
    // a result without the formula's value (`values: 'preserve'`)
    if (result && !res && c.held === undefined) issues.add('noCachedValues', { key: mode === 'formulas' ? 'formulaNoValue' : 'formulaEmpty' }, where(addr))
    const tree = mode === 'formulas' ? fileTree(c.formula, ctx, ctx.tableAt(addr)) : null
    if (mode === 'formulas' && !tree) issues.add('formula', { key: 'formulaAsValue' }, where(addr))
    if (!tree) {
      put(addr, cellXml(addr, c, value))
      continue
    }
    formulas = true
    if (ctx.totalsAt(addr) && c.array === undefined) {
      put(addr, cellXml(addr, c, value, `<f>${escapeText(toFile(tree))}</f>`, false))
      continue
    }
    const ref = c.array ?? res?.spill ?? addr
    put(addr, cellXml(addr, c, value, `<f t="array" ref="${ref}">${escapeText(toFile(tree))}</f>`, c.array === undefined))
  }
  // the cells arrays spilled into, with the look their own cells give them
  const ownAt = (addr: string) => {
    if (!layer) return sheet.cells[addr]
    const { row, col } = parseAddress(addr)
    return inputIn(sheet, addr, row, col)
  }
  if (result)
    for (const addr of Object.keys(result)) {
      const r = result[addr]
      if (r.spillOf !== undefined) put(addr, cellXml(addr, ownAt(addr), r.value))
      if (r.spill !== undefined && r.spillFill !== undefined && mode === 'values')
        for (const a of rectCells(parseRange(r.spill)!)) if (!result[a]) put(a, cellXml(a, ownAt(a), r.spillFill))
    }
  // the collapsed marks as Excel keeps and saves them, whatever marks the book holds
  const rowLines = lines(normalized(sheet.rows, 'r', sheet.outline?.summaryBelow !== false), 'r')
  const colLines = lines(normalized(sheet.columns, 'c', sheet.outline?.summaryRight !== false), 'c')
  for (const [lo, hi] of rowLines) for (let r = lo; r <= hi; r++) rowsUsed.add(r)
  const rowAttrs = new Map<number, string>()
  for (const [lo, hi, e] of rowLines) {
    const xf = e.style !== undefined || e.numberFormat !== undefined ? styles.xf(e.style, e.numberFormat) : 0
    // a height Excel fitted is written without customHeight, as Excel writes it
    const attrs = (e.height !== undefined ? ` ht="${e.height}"${e.bestFit ? '' : ' customHeight="1"'}` : '') + (e.hidden ? ' hidden="1"' : '') + (xf ? ` s="${xf}" customFormat="1"` : '') + outlineAttrs(e)
    for (let r = lo; r <= hi; r++) rowAttrs.set(r, attrs)
  }
  crossings(rowLines, colLines, styles, (r, c, xf) => {
    if (xmlAt[r]?.[c] === undefined && !(layer && layer.idAt(r, c) !== NONE)) put(formatAddress(r, c), `<c r="${formatAddress(r, c)}" s="${xf}"/>`)
  })
  // the rows of the looks' blocks too, their cells written as the rows are, not held cell by cell
  if (layer) for (let b = 0; b < layer.blocks; b++) for (let r = layer.r1[b]; r <= layer.r2[b]; r++) rowsUsed.add(r)
  const data: string[] = []
  for (const row of [...rowsUsed].sort((a, b) => a - b)) {
    let cells = ''
    const cols = colsOf[row] ?? []
    if (cols.some((c, i) => i > 0 && c < cols[i - 1])) cols.sort((a, b) => a - b)
    if (layer) cells = rowXml(row, layer, layerXf!, cols, xmlAt[row])
    else for (const c of cols) cells += xmlAt[row][c]
    data.push(cells ? `<row r="${row + 1}"${rowAttrs.get(row) ?? ''}>${cells}</row>` : `<row r="${row + 1}"${rowAttrs.get(row) ?? ''}/>`)
  }
  const looked = layer?.bounds()
  const any = extent.n > 0 || !!looked
  const [rowLo, rowHi] = [Math.min(extent.r1, looked?.r1 ?? Infinity), Math.max(extent.r2, looked?.r2 ?? -1)]
  const minCol = Math.min(extent.c1, looked?.c1 ?? MAX_COLS)
  const maxCol = Math.max(extent.c2, looked?.c2 ?? 0)
  const dim = any ? `${indexToCol(minCol)}${rowLo + 1}:${indexToCol(maxCol)}${rowHi + 1}` : 'A1'
  const cols = colLines.map(([lo, hi, e]) => {
    const xf = e.style !== undefined || e.numberFormat !== undefined ? styles.xf(e.style, e.numberFormat) : 0
    const size = e.width !== undefined ? ` width="${e.width}" customWidth="1"${e.bestFit ? ' bestFit="1"' : ''}` : ` width="${DEFAULT_WIDTH}"`
    return `<col min="${lo + 1}" max="${hi + 1}"${size}${e.hidden ? ' hidden="1"' : ''}${xf ? ` style="${xf}"` : ''}${outlineAttrs(e)}/>`
  })
  const merges = sheet.merges ?? []
  const used = any ? { r1: rowLo, r2: rowHi, c1: minCol, c2: maxCol } : undefined
  let filter = sheet.filtered && !ctx.filterTable ? filterRange(used, rowLines) : undefined
  // an autofilter of the sheet cannot meet a table: only tables without a header row (no filter of their own) are left
  const area = filter && parseRange(filter)!
  if (area && (sheet.tables ?? []).some((t) => meets(parseRange(t.range)!, area))) {
    issues.add('autoFilter', { key: 'filterModeOut' }, { sheet: sheet.name })
    filter = undefined
  }
  // relationships past the tables': the notes' comments and drawing, then each address of a hyperlink once
  const rels: SheetRel[] = []
  const relId = (r: SheetRel) => `rId${ctx.tableParts.length + rels.push(r)}`
  let drawing: string | undefined
  if (notes.size) {
    relId({ type: 'comments', target: `../comments${index + 1}.xml` })
    drawing = relId({ type: 'vmlDrawing', target: `../drawings/vmlDrawing${index + 1}.vml` })
  }
  const urls = new Map<string, string>()
  const hyperlinks = hyperlinksXml(links, (url) => urls.get(url) ?? urls.set(url, relId({ type: 'hyperlink', target: url })).get(url)!, sheet.name, issues)
  const pane = paneXml(sheet.frozen)
  const names: FileNames = { sheets: ctx.sheets, tables: ctx.tables, externalBooks: ctx.externalBooks }
  const cond = conditionalXml(sheet, index, { styles, issues, ...names })
  // the outline: where its summaries are, and its deepest levels, which Excel reserves room for
  const o = sheet.outline
  const outlinePr = o && (o.summaryBelow !== undefined || o.summaryRight !== undefined) ? `<outlinePr${o.summaryBelow !== undefined ? ` summaryBelow="${o.summaryBelow ? 1 : 0}"` : ''}${o.summaryRight !== undefined ? ` summaryRight="${o.summaryRight ? 1 : 0}"` : ''}/>` : ''
  const deepest = (ls: Line<{ outlineLevel?: number }>[], attr: string) => {
    const n = ls.reduce((m, [, , e]) => Math.max(m, e.outlineLevel ?? 0), 0)
    return n ? ` ${attr}="${n}"` : ''
  }
  const xml =
    `${HEAD}<worksheet xmlns="${MAIN}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    (filter || outlinePr ? `<sheetPr${filter ? ' filterMode="1"' : ''}${outlinePr ? `>${outlinePr}</sheetPr>` : '/>'}` : '') +
    `<dimension ref="${dim}"/>${sheetViewXml(sheet.view, pane, ctx.active)}` +
    // without it Excel for Mac takes a standard width of 10 digits, not the 8 the model assumes
    `<sheetFormatPr defaultRowHeight="15"${deepest(rowLines, 'outlineLevelRow')}${deepest(colLines, 'outlineLevelCol')}/>` +
    (cols.length ? `<cols>${cols.join('')}</cols>` : '') +
    `<sheetData>${data.join('')}</sheetData>` +
    (filter ? `<autoFilter ref="${filter}">${KEEP_ALL}</autoFilter>` : '') +
    (merges.length ? `<mergeCells count="${merges.length}">${merges.map((m) => `<mergeCell ref="${escapeAttr(m)}"/>`).join('')}</mergeCells>` : '') +
    cond.main +
    validationsXml(sheet.validations, { sheet: sheet.name, ...names, issues }) +
    hyperlinks +
    (drawing ? `<legacyDrawing r:id="${drawing}"/>` : '') +
    (ctx.tableParts.length ? `<tableParts count="${ctx.tableParts.length}">${ctx.tableParts.map((id) => `<tablePart r:id="${id}"/>`).join('')}</tableParts>` : '') +
    (cond.ext ? `<extLst>${cond.ext}</extLst>` : '') +
    '</worksheet>'
  return { xml, formulas, rels, ...(notes.size ? { notes: { comments: commentsXml(notes, sheet.name, issues), vml: vmlXml(notes, ctx.noteBlock), count: notes.size } } : {}) }
}

/**
 * The cell format of each look of the sheet's looks (`sAttr` of a cell of it, ` s="n"` or none), numbered with the sheet's
 * entries as entries of theirs would be: each cell of the looks just before the first entry after it by address, the
 * entries in their order, as a file read lists them (a cell a note or a link alone made comes last); by cell where the
 * sheet's rows or columns have a look, which a cell's whole look takes.
 */
function numberLooks(sheet: SheetInput, layer: Layer, sAttr: (addr: string, c: CellInput | undefined) => string): (r: number, c: number, look: number) => string {
  const lined = [sheet.rows, sheet.columns].some((m) => m && Object.values(m).some((e) => e.style !== undefined || e.numberFormat !== undefined))
  const byLook: (string | undefined)[] = []
  const attr = (r: number, c: number, look: number) => (lined ? sAttr(formatAddress(r, c), layer.looks.list[look]) : (byLook[look] ??= sAttr(formatAddress(r, c), layer.looks.list[look])))
  // the looks not numbered yet; once none is left, only the entries take new numbers
  let left = lined ? Infinity : layer.used().length
  // a cursor through the looks' cells: block, row, run and column; a run's first cell alone without lines
  let [b, k] = [0, 0]
  let r = layer.blocks ? layer.r1[0] : 0
  let c = layer.blocks ? layer.patterns[layer.at[0]][0] : 0
  const before = (place: number) => {
    while (left > 0 && b < layer.blocks) {
      const p = layer.patterns[layer.at[b]]
      if (r * MAX_COLS + c >= place) return
      if (!lined && byLook[p[k + 2]] === undefined) left--
      attr(r, c, p[k + 2])
      if (lined && c < p[k + 1]) c++
      else if ((k += 3) < p.length) c = p[k]
      else if (r < layer.r2[b]) [r, k, c] = [r + 1, 0, p[0]]
      else if (++b < layer.blocks) [r, k, c] = [layer.r1[b], 0, layer.patterns[layer.at[b]][0]]
    }
  }
  const at = { r: 0, c: 0 }
  for (const a in sheet.cells) {
    readAddress(a, at)
    before(at.r * MAX_COLS + at.c)
    sAttr(a, sheet.cells[a])
  }
  before(Infinity)
  return attr
}

/** A row's cells: the entries' (their columns `cols` in order, their XML by column in `at`) and the looks' between them, in order of columns. */
function rowXml(row: number, layer: Layer, attr: (r: number, c: number, look: number) => string, cols: readonly number[], at: readonly string[] | undefined): string {
  const p = layer.row(row)
  let xml = ''
  let i = 0
  for (let k = 0; k < p.length; k += 3)
    for (let c = p[k]; c <= p[k + 1]; c++) {
      for (; i < cols.length && cols[i] < c; i++) xml += at![cols[i]]
      // an entry's own cell (a spilled one, a crossing) is written as it is
      if (i < cols.length && cols[i] === c) continue
      xml += `<c r="${formatAddress(row, c)}"${attr(row, c, p[k + 2])}/>`
    }
  for (; i < cols.length; i++) xml += at![cols[i]]
  return xml
}

/**
 * The cells where a row with a look crosses a column with one and Excel would show another cell format than the model:
 * Excel shows such a cell without one of its own by the row's whole cell format (by the column's when the row's is the
 * default and so not written), the model each of its style and format from the row, else the column. Excel writes these
 * cells itself where lines it formats cross; `put` gets each with the model's cell format.
 */
function crossings(rows: Line<RowInput>[], cols: Line<ColumnInput>[], styles: StyleWriter, put: (r: number, c: number, xf: number) => void) {
  const looked = <T extends Look>(ls: Line<T>[]) => ls.filter(([, , e]) => e.style !== undefined || e.numberFormat !== undefined)
  const withCols = looked(cols)
  if (!withCols.length) return
  for (const [rlo, rhi, row] of looked(rows)) {
    const rowXf = styles.xf(row.style, row.numberFormat)
    for (const [clo, chi, col] of withCols) {
      const look = lookOf(undefined, row, col)
      const xf = styles.xf(look.style, look.numberFormat)
      if (xf === (rowXf || styles.xf(col.style, col.numberFormat))) continue
      for (let r = rlo; r <= rhi; r++) for (let c = clo; c <= chi; c++) put(r, c, xf)
    }
  }
}

/** A row's or a column's outline level and collapsed mark as attributes. */
const outlineAttrs = (e: { outlineLevel?: number; collapsed?: boolean }) => (e.outlineLevel ? ` outlineLevel="${e.outlineLevel}"` : '') + (e.collapsed ? ' collapsed="1"' : '')

/** The written cells' rectangle (`cells`: rows and columns, or none) stretched over the hidden rows and the row above the first. */
function filterRange(cells: { r1: number; r2: number; c1: number; c2: number } | undefined, rowLines: Line<RowInput>[]): string {
  let { r1, r2, c1, c2 } = cells ?? { r1: Infinity, r2: -1, c1: 0, c2: 0 }
  for (const [lo, hi, e] of rowLines)
    if (e.hidden) {
      r1 = Math.min(r1, Math.max(0, lo - 1))
      r2 = Math.max(r2, hi)
    }
  return r2 < 0 ? 'A1' : `${formatAddress(r1, c1)}:${formatAddress(r2, c2)}`
}
