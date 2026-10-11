/**
 * Structure operations against Excel en-US: each case is a small book and one operation (`StructureOp`);
 * Excel applies the operation to the book and the formulas of every sheet and the defined names are read
 * back (`excel-struct.applescript`). `npx tsx bench/coverage/run.ts struct <set>` runs the cases of
 * `struct/<set>.ts` and stores the answers in `test/excel/struct/<set>.json`;
 * `test/structure-excel.test.ts` compares `editStructure` with them without Excel. A book with styles, columns,
 * rows, merged areas or number formats is written with them, and the look of `LOOK` (and of the set's `look` rectangles) is read back too.
 * `FSTRUCT_ONLY=id,id` runs only those cases, `FCOV_FRESH=1` runs every case again.
 */
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import ExcelJS from 'exceljs'
import { answeredLine, inCaseOrder, keptSame, loadStored, runCases, secondsSince, setModule, took, writeAnswers } from './stored.js'
import { localeMeta } from './locales.js'
import { ensureExcel, excelVersion, runWatched } from './excel.js'
import { markDynamic } from './book.js'
import { formatCode, formulaValue, writeNames, writeParts, type FileParts } from './case-book.js'
import type { BookInput, CellStyle, NumberFormat, RowInput, SheetInput } from '../../src/contract.js'
import { lines } from '../../src/layout/spans.js'
import { formatAddress, indexToCol, parseAddress } from '../../src/model.js'
import type { StructCase, StructCases, StructSet } from './struct/cases.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
const SCRIPT = join(HERE, 'excel-struct.applescript')

/** The rectangle read on every sheet after the operation; a case keeps its cells inside it or lists them in `extra`. */
export const READ = 'A1:Z60'

/** The rectangle whose look is read on every sheet of a styled book; a case keeps its styles inside it. */
export const LOOK = 'A1:J24'

/** What a cell looks like in the short signature: bold, the fill colour, the number format; left out when plain. */
export interface CellLook {
  bold?: true
  fill?: string
  format?: string
}

/** The look of `LOOK` on a sheet: cells that are not plain, the columns and rows not of the standard size, merged areas meeting it, hidden columns and rows. */
export interface SheetLook {
  cells: Record<string, CellLook>
  widths: Record<string, number>
  heights: Record<string, number>
  merges: string[]
  /** Hidden columns and rows (`C`, `5`) when there are any. */
  hidden?: string[]
}

export interface ExcelStructAnswer {
  /** Excel's error when it refused the operation (AppleScript error number and text). */
  error?: string
  /** What Excel shows in each non-empty cell of `READ` (`=SUM(A1:A3)`, `5`), the sheet's used range, and the look of a styled book. */
  sheets: { name: string; used: string; cells: Record<string, string>; look?: SheetLook }[]
  extra: { sheet: string; cell: string; text: string }[]
  /** `name` is `Sheet!name` for a local name; `ref` as Excel shows it (`=S!$A$5`). */
  names: { name: string; ref: string }[]
}

export interface StoredStruct extends StructCases {
  meta: Record<string, string | number>
  answers: Record<string, ExcelStructAnswer>
}

/** A book has styles to write and a look to read. */
export const styled = (book: BookInput) =>
  !!book.styles?.length || book.sheets.some((s) => s.columns || s.rows || s.merges || Object.values(s.cells).some((c) => c.style !== undefined || c.format !== undefined))

const argb = (color: string) => ({ argb: `FF${color.slice(1).toUpperCase()}` })

/** A style and a number format as ExcelJS takes them; a font keeps the default Calibri 11. */
function excelStyle(st: CellStyle | undefined, format: NumberFormat | undefined): Partial<ExcelJS.Style> {
  const out: Partial<ExcelJS.Style> = {}
  if (format !== undefined) out.numFmt = formatCode(format)
  if (st?.font) {
    const { color, ...f } = st.font
    out.font = { name: 'Calibri', size: 11, ...f, ...(color ? { color: argb(color) } : {}) }
  }
  if (st?.fill) out.fill = { type: 'pattern', pattern: 'solid', fgColor: argb(st.fill.color) }
  if (st?.border) out.border = Object.fromEntries(Object.entries(st.border).map(([side, l]) => [side, { style: l.style, ...(l.color ? { color: argb(l.color) } : {}) }]))
  if (st?.alignment) {
    // ExcelJS spells xlsx's vertical center `middle` and has no `general` (the default)
    const { wrap, vertical, horizontal, ...a } = st.alignment
    out.alignment = { ...a, ...(wrap !== undefined ? { wrapText: wrap } : {}), ...(vertical ? { vertical: vertical === 'center' ? 'middle' : vertical } : {}), ...(horizontal && horizontal !== 'general' ? { horizontal } : {}) }
  }
  return out
}

/**
 * Columns, rows and the cells' own style and format of a sheet. ExcelJS writes a column with a style and
 * no width as width 9, and drops a row with a style and neither a height nor cells: cases avoid both.
 */
function writeLayout(c: StructCase, book: BookInput, s: SheetInput, ws: ExcelJS.Worksheet) {
  const style = (i: number | undefined) => (i === undefined ? undefined : book.styles![i])
  const cols = lines(s.columns, 'c')
  const rows = lines(s.rows, 'r')
  for (const [lo, hi, e] of cols) {
    if (e.width === undefined) throw new Error(`${c.id}: column ${indexToCol(lo)} has no width (ExcelJS would write 9)`)
    for (let i = lo; i <= hi; i++) Object.assign(ws.getColumn(i + 1), { width: e.width, style: excelStyle(style(e.style), e.format), ...(e.hidden ? { hidden: true } : {}) })
  }
  const filled = new Set(Object.keys(s.cells).map((a) => parseAddress(a).row))
  for (const [lo, hi, e] of rows)
    for (let i = lo; i <= hi; i++) {
      // a row that is only hidden is written by `writeParts`
      if (e.height === undefined && !filled.has(i) && Object.keys(e).some((k) => k !== 'hidden')) throw new Error(`${c.id}: row ${i + 1} has neither a height nor cells (ExcelJS would drop it)`)
      // the row's own style, without passing it to its cells as ExcelJS's setters do
      Object.assign(ws.getRow(i + 1), { style: excelStyle(style(e.style), e.format) }, e.height === undefined ? {} : { height: e.height })
    }
  // Excel takes an empty cell's whole format from its row, else its column, where a cell here takes its style
  // and its number format separately: a row with only one of them crossing a formatted column could part on that
  for (const [lo, , r] of rows)
    if ((r.style === undefined) !== (r.format === undefined) && cols.some(([, , k]) => k.style !== undefined || k.format !== undefined)) throw new Error(`${c.id}: row ${lo + 1} has only a style or a format, and a column is formatted`)
  // a cell without its own style or format shows its row's, else its column's
  for (const [addr, cell] of Object.entries(s.cells)) {
    const { row, col } = parseAddress(addr)
    const r = rows.find(([lo, hi]) => row >= lo && row <= hi)?.[2]
    const k = cols.find(([lo, hi]) => col >= lo && col <= hi)?.[2]
    ws.getCell(addr).style = excelStyle(style(cell.style ?? r?.style ?? k?.style), cell.format ?? r?.format ?? k?.format)
  }
  // ExcelJS gives every cell of a merged area the style of its top-left cell: cases keep them equal
  const look = (addr: string) => JSON.stringify(ws.getCell(addr).style ?? {})
  for (const m of s.merges ?? []) {
    const [a, b] = m.split(':').map(parseAddress)
    for (let r = a.row; r <= b.row; r++) for (let k = a.col; k <= b.col; k++) if (look(formatAddress(r, k)) !== look(formatAddress(a.row, a.col))) throw new Error(`${c.id}: ${formatAddress(r, k)} of the merged area ${m} has a style of its own`)
    ws.mergeCells(m)
  }
}

/** The file parts of a book: its hidden rows and rows of height 0, and the autofilter `files` gives each `filtered` sheet. */
function partsOf(c: StructCase, book: BookInput, files: FileParts = {}): FileParts {
  const sheets: FileParts['sheets'] = {}
  for (const s of book.sheets) {
    const spans = (is: (e: RowInput) => boolean) => lines(s.rows, 'r').flatMap(([lo, hi, e]) => (is(e) ? [lo === hi ? `${lo + 1}` : `${lo + 1}:${hi + 1}`] : []))
    const hiddenRows = spans((e) => !!e.hidden)
    const zeroRows = spans((e) => e.height === 0)
    const filter = files.sheets?.[s.name]?.filter
    if (!!s.filtered !== !!filter) throw new Error(`${c.id}: sheet ${s.name} is filtered without a filter in the file, or the other way round`)
    if (hiddenRows.length || zeroRows.length || filter) sheets[s.name] = { hiddenRows, ...(zeroRows.length ? { zeroRows } : {}), ...(filter ? { filter } : {}) }
  }
  return { sheets }
}

/**
 * The book as xlsx: values and formulas (no cached results; spill references as Excel stores them), names with
 * their scope, the layout of a styled book with its hidden lines and filters; with `dynamic`, formulas as dynamic
 * array formulas (Excel 365).
 */
export async function writeBook(c: StructCase, book: BookInput, file: string, dynamic: boolean, files?: FileParts): Promise<void> {
  const wb = new ExcelJS.Workbook()
  for (const s of book.sheets) {
    const ws = wb.addWorksheet(s.name)
    for (const [addr, cell] of Object.entries(s.cells)) {
      if (cell.value !== null && typeof cell.value === 'object') throw new Error(`${c.id}: dates are not used in structure cases`)
      // an old-kind array formula over its area, else a dynamic one in a dynamic set
      const formula = cell.formula === undefined ? undefined : formulaValue(cell.formula, cell.array ?? (dynamic ? addr : undefined))
      ws.getCell(addr).value = (formula ?? cell.value ?? null) as ExcelJS.CellValue
    }
    if (styled(book)) writeLayout(c, book, s, ws)
  }
  writeNames(wb, book)
  await wb.xlsx.writeFile(file)
  if (dynamic) await markDynamic(file, book.sheets.length, (i, addr) => book.sheets[i].cells[addr]?.array !== undefined)
  const parts = partsOf(c, book, files)
  if (Object.keys(parts.sheets!).length) await writeParts(file, book, parts)
}

/**
 * The AppleScript arguments of an operation: sheet, range, target sheet or new name, target range. Excel's
 * `range` takes cell ranges only: whole columns and rows are spelled as cells (`B1:B1048576`), rows and
 * columns to insert or delete as the cells of their first column or row.
 */
export function opArgs(c: StructCase): string[] {
  const o = c.op
  const cells = (r: string) => r.replace(/^([A-Z]+):([A-Z]+)$/, (_, a, b) => `${a}1:${b}1048576`).replace(/^(\d+):(\d+)$/, (_, a, b) => `A${a}:XFD${b}`)
  switch (o.op) {
    case 'insertRows':
    case 'deleteRows':
      return [o.sheet, `A${o.row}:A${o.row + (o.count ?? 1) - 1}`, '', '']
    case 'insertColumns':
    case 'deleteColumns':
      return [o.sheet, `${o.column}1:${colAt(o.column, (o.count ?? 1) - 1)}1`, '', '']
    case 'moveRange':
      return [o.from.sheet, cells(o.from.range), o.to.sheet, o.to.cell]
    case 'copyRange':
      return [o.from.sheet, cells(o.from.range), o.to.sheet, cells(o.to.range)]
    case 'renameSheet':
      return [o.sheet, '', o.name, '']
    case 'deleteSheet':
      return [o.sheet, '', '', '']
    case 'insertSheet':
      return [o.before ?? '', '', o.name, '']
    case 'moveSheet':
      return [o.sheet, '', o.before ?? '', '']
    default:
      throw new Error(`${c.id}: the struct step has no ${o.op} (the tables step has)`)
  }
}

/** A case's operation as the scripts of the steps with files take it: JSON of the operation and its arguments, `[]` for none. */
export const opJson = (c: { id: string; book: string; op?: StructCase['op'] }): string => JSON.stringify(c.op ? [c.op.op, ...opArgs({ id: c.id, book: c.book, op: c.op })] : [])

function colAt(letters: string, plus: number): string {
  let n = 0
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64)
  n += plus
  let s = ''
  while (n > 0) {
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

/** Formulas of `READ` as Excel returns them (a list of rows) by address, empty cells left out. */
export function cellsOf(rows: unknown): Record<string, string> {
  const out: Record<string, string> = {}
  const grid = Array.isArray(rows) ? (rows as unknown[][]) : [[rows]]
  const [, c0, r0] = /^([A-Z]+)(\d+)/.exec(READ)!
  grid.forEach((row, i) =>
    row.forEach((v, j) => {
      const text = String(v)
      if (text !== '') out[`${colAt(c0, j)}${Number(r0) + i}`] = text
    }),
  )
  return out
}

/** The cells, columns and rows of the look's rectangles (`LOOK,…`) as AppleScript arguments (lines). */
export function lookArgs(look: string): [string, string, string] {
  const [cells, cols, rows] = [new Set<string>(), new Set<string>(), new Set<string>()]
  for (const rect of look.split(',')) {
    const [a, b] = rect.split(':').map(parseAddress)
    for (let r = a.row; r <= b.row; r++) for (let c = a.col; c <= b.col; c++) cells.add(formatAddress(r, c))
    for (let c = a.col; c <= b.col; c++) cols.add(indexToCol(c))
    for (let r = a.row; r <= b.row; r++) rows.add(String(r + 1))
  }
  return [[...cells].join('\n'), [...cols].join('\n'), [...rows].join('\n')]
}

export type RawLook = [[string, boolean, number[], string, string][], [string, boolean, number, boolean][], [string, boolean, number, boolean][]]

/** Excel's look of a sheet: a white fill is no fill (cases do not use white), sizes of the standard size are left out. */
export function lookOf([cells, widths, heights]: RawLook): SheetLook {
  const hex = (rgb: number[]) => `#${rgb.map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase()}`
  const out: SheetLook = { cells: {}, widths: {}, heights: {}, merges: [] }
  for (const [addr, bold, rgb, format, merge] of cells) {
    const fill = hex(rgb)
    const look: CellLook = { ...(bold ? { bold: true } : {}), ...(fill !== '#FFFFFF' ? { fill } : {}), ...(format !== 'General' ? { format } : {}) }
    if (Object.keys(look).length) out.cells[addr] = look
    const m = merge.replaceAll('$', '')
    if (m && !out.merges.includes(m)) out.merges.push(m)
  }
  const hidden: string[] = []
  for (const [col, std, w, hid] of widths) if (hid) hidden.push(col)
  else if (!std) out.widths[col] = w
  for (const [row, std, h, hid] of heights) if (hid) hidden.push(row)
  else if (!std) out.heights[row] = h
  return hidden.length ? { ...out, hidden } : out
}

/** One case in Excel. A window that is not a book stops the run (nothing is clicked); the repair window is declined. */
async function runCase(c: StructCase, book: BookInput, dynamic: boolean, lookRects: string, files?: FileParts): Promise<ExcelStructAnswer> {
  const extra = (c.extra ?? []).map((e) => `${e.sheet}\t${e.cell}`).join('\n')
  const look = styled(book) ? lookArgs(lookRects) : ['', '', '']
  return runWatched<ExcelStructAnswer>({
    script: SCRIPT,
    label: c.id,
    folder: 'fstruct',
    empty: true,
    book: (file) => writeBook(c, book, file, dynamic, files),
    args: () => [c.via ?? c.op.op, ...opArgs(c), READ, extra, ...look],
    timeout: 90_000,
    poll: 1000,
    after: 'case',
    fatal: true,
    // the run stops on a book Excel does not load
    repair: async (stop) => {
      await stop()
      throw new Error(`${c.id}: Excel did not accept the book (repair window declined)`)
    },
    done: (out) => {
      const raw = JSON.parse(out) as { error: string; sheets: [string, string, unknown, RawLook | ''][]; extra: [string, string, string][]; names: [string, string][] }
      return {
        ...(raw.error ? { error: raw.error } : {}),
        sheets: raw.sheets.map(([name, used, f, look]) => ({ name, used, cells: cellsOf(f), ...(look ? { look: lookOf(look) } : {}) })),
        extra: raw.extra.map(([sheet, cell, text]) => ({ sheet, cell, text })),
        names: raw.names.map(([name, ref]) => ({ name, ref })),
      }
    },
  })
}

/** A case is the same while it and its book are: its stored key and its current one, by the set each comes from. */
export const structKey = (c: StructCase, set: Pick<StructCases, 'books' | 'files'>) => JSON.stringify([c, set.books[c.book], set.files?.[c.book]])

/** Runs the set's cases that have no stored answer for the same case (FCOV_FRESH=1 runs all) and stores the answers. */
export async function stepStruct(set: string): Promise<void> {
  const built = (await setModule<StructSet>('struct', set)).buildStructCases()
  const { books, cases, dynamic = false, look = [], widths, files } = built
  const lookRects = [LOOK, ...look].join(',')
  const target = join(ROOT, `test/excel/struct/${set}.json`)
  const answers = keptSame(built, loadStored<StoredStruct>(target), structKey)
  const t0 = Date.now()
  const save = () => {
    const meta = { version: excelVersion(), locale: localeMeta('en-US'), read: READ, ...(Object.values(books).some(styled) ? { look: lookRects } : {}), seconds: secondsSince(t0) }
    writeAnswers(target, { meta, books, cases, ...(dynamic ? { dynamic } : {}), ...(widths === false ? { widths } : {}), ...(files ? { files } : {}), answers: inCaseOrder(cases, answers) })
  }
  await ensureExcel('en-US')
  await runCases(cases.filter((c) => !answers[c.id]), 'FSTRUCT_ONLY', async (c) => void (answers[c.id] = await runCase(c, books[c.book], dynamic, lookRects, files?.[c.book])), (done) => {
    save()
    console.log(`${answeredLine(`struct ${set}`, done, answers, cases)}, ${took(t0)} s`)
  })
}
