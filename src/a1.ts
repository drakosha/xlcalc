/**
 * A1 addresses without the engine (the `./a1` subpath, the same functions in the root): cells, ranges, the keys of layout
 * lines and their maps patched, a sheet's outline groups, rectangles, a sheet's name in a reference as Excel writes it
 * and the rule names compare by. Rows and columns are numbered from 1, as A1 numbers them; the parts reused are the
 * engine's own (`model.ts`, `structure/area.ts`, `layout/spans.ts`, `layout/outline.ts`, `collate.ts`), which compute
 * nothing. Each function refuses options it does not take (`strict-args.ts`).
 */
import type { CellPosition, CellRect, LinePatch, LineSpan, OutlineGroup, RangeRect, SheetInput, SheetOutline } from './contract.js'
import { caseKey } from './collate.js'
import { InputError, what } from './input-error.js'
import { groupsOf, type OutlineEntry } from './layout/outline.js'
import { lines, parseSpan, patchLines as patchSpans, spanKey } from './layout/spans.js'
import type { Locale } from './locale/types.js'
import { colToIndex, formatAddress, indexToCol, MAX_COLS, MAX_ROWS } from './model.js'
import { contains, encode, grown, meets, parseRange as rangeArg, span, type Area, type Axis, type Rect } from './structure/area.js'
import { PLAIN_SHEET } from './sheet-pattern.js'
import { strict } from './strict-args.js'

export { MAX_COLS, MAX_ROWS }

/** A whole number from 1 to `max`, else `InputError` naming `path`. */
function line(n: unknown, max: number, path: string): number {
  if (typeof n !== 'number' || !Number.isInteger(n) || n < 1 || n > max) throw new InputError(path, `expected an integer from 1 to ${max}, got ${what(n)}`)
  return n
}

const text = (v: unknown, path: string): string => {
  if (typeof v !== 'string') throw new InputError(path, `expected a string, got ${what(v)}`)
  return v
}

/** Column letters of column `n` (1 is `A`, 16,384 is `XFD`). */
export const colLetters = /* @__PURE__ */ strict(function colLetters(n: number): string {
  return indexToCol(line(n, MAX_COLS, 'n') - 1)
})

/** The number of a column by its letters in any case (`B` is 2), or null when they are no column of the sheet (`XFE`). */
export const colNumber = /* @__PURE__ */ strict(function colNumber(letters: string): number | null {
  if (!/^[A-Za-z]{1,3}$/.test(text(letters, 'letters'))) return null
  const n = colToIndex(letters) + 1
  return n <= MAX_COLS ? n : null
})

/** The address of a cell (`B12`). */
export const cellAddress = /* @__PURE__ */ strict(function cellAddress(row: number, col: number): string {
  return formatAddress(line(row, MAX_ROWS, 'row') - 1, line(col, MAX_COLS, 'col') - 1)
})

/** `$` where a reference may hold it: at the start, after `:`, between a column's letters and a row's number. */
const DOLLAR = /(^|:)\$|([A-Za-z])\$(?=\d)/g
const COLUMNS = /^[A-Z]+:[A-Z]+$/
const ROWS = /^\d+:\d+$/

/**
 * A range's rectangle and kind: `B2:D9` and `B2` are cells, `B:D` whole columns, `2:9` whole rows, with `$` and letters
 * in any case; corners given the other way round are put in order, as Excel orders a reference typed so. Null for
 * anything else or a range that leaves the sheet.
 */
export const parseRange = /* @__PURE__ */ strict(function parseRange(range: string): RangeRect | null {
  // letters of other scripts stay out: some come to Latin ones in upper case (`ß` is `SS`)
  if (!/^[A-Za-z0-9$:]+$/.test(text(range, 'range'))) return null
  const t = range.replace(DOLLAR, '$1$2').toUpperCase()
  const x = rangeArg(t)
  if (!x) return null
  const kind = COLUMNS.test(t) ? 'columns' : ROWS.test(t) ? 'rows' : 'cells'
  return { kind, top: x.r1 + 1, left: x.c1 + 1, bottom: x.r2 + 1, right: x.c2 + 1 }
})

/** A cell's row and column (`B12`, `$B$12`, `b12`), or null when the text is no cell of the sheet (`ZZZ1`, `A0`, `B2:C3`). */
export const parseCell = /* @__PURE__ */ strict(function parseCell(cell: string): CellPosition | null {
  const r = text(cell, 'cell').includes(':') ? null : parseRange(cell)
  return r?.kind === 'cells' ? { row: r.top, col: r.left } : null
})

/** A rectangle checked against the sheet, corners in order, as the engine's area of a reference without `$`. */
function areaOf(r: CellRect, path = 'rect'): Area {
  const [top, bottom] = [line(r.top, MAX_ROWS, `${path}.top`), line(r.bottom, MAX_ROWS, `${path}.bottom`)]
  const [left, right] = [line(r.left, MAX_COLS, `${path}.left`), line(r.right, MAX_COLS, `${path}.right`)]
  const b = (i: number) => ({ i: i - 1, abs: false })
  return { kind: top === bottom && left === right ? 'cell' : 'range', r: span(b(top), b(bottom)), c: span(b(left), b(right)) }
}

/** A rectangle's address as Excel writes it: `B2` for one cell, `B2:D9`, `B:D` over every row, `2:9` over every column. */
export const rangeAddress = /* @__PURE__ */ strict(function rangeAddress(rect: CellRect): string {
  return encode(areaOf(rect))
})

const rect = (x: CellRect): Rect => ({ r1: x.top, c1: x.left, r2: x.bottom, c2: x.right })

/** Whether two rectangles share a cell. */
export const rectsMeet = /* @__PURE__ */ strict(function rectsMeet(a: CellRect, b: CellRect): boolean {
  return meets(rect(a), rect(b))
})

/** Whether `inner` lies within `outer`. */
export const rectContains = /* @__PURE__ */ strict(function rectContains(outer: CellRect, inner: CellRect): boolean {
  return contains(rect(outer), rect(inner))
})

/**
 * The rectangle a merge takes in (`mergeCells`, Merge Cells): `rect` grown by each merged area of `merges` it meets,
 * taken in whole, again for each it then meets, until it meets no other.
 */
export const mergedArea = /* @__PURE__ */ strict(function mergedArea(r: CellRect, merges: readonly CellRect[]): CellRect {
  areaOf(r)
  if (!Array.isArray(merges)) throw new InputError('merges', `expected an array of rectangles, got ${what(merges)}`)
  merges.forEach((m, i) => areaOf(m, `merges[${i}]`))
  const g = grown(rect(r), merges.map(rect))
  return { top: g.r1, left: g.c1, bottom: g.r2, right: g.c2 }
})

/** The engine's axis of layout lines. */
const axisOf = (axis: 'columns' | 'rows') => {
  if (axis !== 'columns' && axis !== 'rows') throw new InputError('axis', `expected "columns" or "rows", got ${what(axis)}`)
  return axis === 'columns' ? 'c' : 'r'
}

/**
 * The lines a key of `SheetInput.columns` or `rows` names (`C`, `B:D`; `5`, `5:9`), or null for a key a book does not
 * take: lower case, a span of one line (`C:C`) or the other way round, past the sheet.
 */
export const parseLines = /* @__PURE__ */ strict(function parseLines(key: string, axis: 'columns' | 'rows'): LineSpan | null {
  const s = parseSpan(text(key, 'key'), axisOf(axis))
  return s && { first: s[0] + 1, last: s[1] + 1 }
})

/** The key of lines `first` to `last` (`C`, `B:D`; `5`, `5:9`), in order whichever comes first. */
export const linesKey = /* @__PURE__ */ strict(function linesKey(first: number, last: number, axis: 'columns' | 'rows'): string {
  const ax = axisOf(axis)
  const max = ax === 'c' ? MAX_COLS : MAX_ROWS
  const [a, b] = [line(first, max, 'first'), line(last, max, 'last')]
  return spanKey(Math.min(a, b) - 1, Math.max(a, b) - 1, ax)
})

const isObject = (v: unknown): v is Record<string, unknown> => typeof v === 'object' && v !== null && !Array.isArray(v)

/**
 * A map of `SheetInput.columns` or `rows` with `patch` applied to lines `span` (`C`, `B:D`; `5`, `5:9`, the axis by the
 * key's shape), as `setLayout` applies it: a value sets a field, null removes it, a field left out stays; spans the patch
 * cuts are split, lines left without a field are dropped. A new map; the one given is not changed. The map is checked as
 * a book's: an object under each span of the axis, no two spans overlapping.
 */
export const patchLines = /* @__PURE__ */ strict(function patchLines<T extends object>(map: Record<string, T> | undefined, span: string, patch: LinePatch<T>): Record<string, T> {
  const axis: Axis = /^[A-Z]/.test(text(span, 'span')) ? 'c' : 'r'
  const s = parseSpan(span, axis)
  if (!s) throw new InputError('span', `expected columns or rows ("B:D", "5:9"), got ${what(span)}`)
  checkMap(map, axis, 'map', `a span of the axis of "${span}"`)
  if (!isObject(patch)) throw new InputError('patch', `expected fields, got ${what(patch)}`)
  return patchSpans(map, axis, s[0], s[1], patch)
})

/** A map of `columns` or `rows` checked as a book's: an object under each span of the axis, no two spans overlapping. */
function checkMap(map: unknown, axis: Axis, path: string, key: string): asserts map is Record<string, object> | undefined {
  if (map !== undefined && !isObject(map)) throw new InputError(path, `expected lines by span, got ${what(map)}`)
  for (const [k, v] of Object.entries(map ?? {})) if (!parseSpan(k, axis) || !isObject(v)) throw new InputError(`${path}[${JSON.stringify(k)}]`, `expected an object under ${key}`)
  if (lines(map as Record<string, object> | undefined, axis).some(([lo], i, ls) => i > 0 && lo <= ls[i - 1][1])) throw new InputError(path, 'two spans overlap')
}

/**
 * A sheet's groups of rows and columns as Excel draws them (`SheetOutline`), from its lines' outline levels, hidden
 * state and collapsed marks and its summary settings: a `SheetInput`, or a `SheetLayout` as `getLayout()` gives it. Each
 * axis's groups by first line, the outer before the inner where two begin together; a group is collapsed when Excel
 * keeps the mark on its summary line (`RowInput.collapsed`), whatever marks the sheet holds.
 */
export const outlineGroups = /* @__PURE__ */ strict(function outlineGroups(sheet: Pick<SheetInput, 'rows' | 'columns' | 'outline'>): SheetOutline {
  if (!isObject(sheet)) throw new InputError('sheet', `expected a sheet or its layout, got ${what(sheet)}`)
  checkMap(sheet.rows, 'r', 'sheet.rows', 'a span of rows ("5", "5:9")')
  checkMap(sheet.columns, 'c', 'sheet.columns', 'a span of columns ("C", "B:D")')
  const o = sheet.outline
  const groups = (map: Record<string, OutlineEntry> | undefined, axis: Axis, after: boolean): OutlineGroup[] =>
    groupsOf(map, axis, after).map((g) => ({ first: g.lo + 1, last: g.hi + 1, level: g.level, summary: g.summary === null ? null : g.summary + 1, collapsed: g.collapsed }))
  return { rows: groups(sheet.rows, 'r', o?.summaryBelow !== false), columns: groups(sheet.columns, 'c', o?.summaryRight !== false) }
})

const PLAIN = new RegExp(`^${PLAIN_SHEET}$`, 'u')
const CELL = /^([A-Za-z]{1,3})(\d+)$/
/** An R1C1 reference at the start of a name: `R` and its number, then `C` and its number; or `C` and its number. */
const R1C1 = /^(?:([Rr])(\d*)(?:[Cc](\d*))?|[Cc](\d*))/

/** A number of a reference (`''` for none) that is a row or column of the sheet. */
const inGrid = (digits: string, max: number) => digits === '' || (Number(digits) >= 1 && Number(digits) <= max)

/**
 * Whether a plain name begins with an R1C1 reference Excel quotes: the whole name (`R`, `RC`, `R1C1`), or one with a number
 * that a `.` does not follow (`R2D2`, `C3PO`, `R1x`; not `C1.x`, `Region`, `R0`, `R1048577`). A column number off the
 * sheet leaves the row's reference (`R1C0`).
 */
function r1c1Start(name: string): boolean {
  const m = R1C1.exec(name)
  if (!m) return false
  const [ref, r, row, col, alone] = m
  if (r === undefined) return inGrid(alone, MAX_COLS) && (ref.length === name.length || (alone !== '' && name[ref.length] !== '.'))
  if (!inGrid(row, MAX_ROWS)) return false
  // `C` and a number off the sheet: the reference is the row's alone
  const end = col === undefined || inGrid(col, MAX_COLS) ? ref.length : 1 + row.length
  const numbered = row !== '' || (end === ref.length && col !== undefined && col !== '')
  return end === name.length || (numbered && name[end] !== '.')
}

/**
 * A sheet's name in a reference as Excel writes it: in quotes, an apostrophe inside doubled, when it is not a plain name
 * (`'Q1 data'`, `'2024'`, `'O''Brien'`), when it is a cell of the sheet (`'A1'`, `'XFD1048576'`; not `XFE1`), when it
 * begins with an R1C1 reference (`'R'`, `'R1C1'`, `'R2D2'`) or when it is one of the locale's booleans (`'TRUE'`;
 * `'ИСТИНА'` with ru-RU). `locale` gives the booleans, en-US's (the stored text, `Range.Formula`) without one. The engine's
 * stored formulas quote every sheet: this is the text a person reads.
 */
export const quoteSheetIfNeeded = /* @__PURE__ */ strict(function quoteSheetIfNeeded(name: string, locale?: Pick<Locale, 'booleans'>): string {
  const n = text(name, 'name')
  const words = locale?.booleans ?? { TRUE: 'TRUE', FALSE: 'FALSE' }
  const cell = CELL.exec(n)
  const quote =
    !PLAIN.test(n) ||
    caseKey(n) === caseKey(words.TRUE) ||
    caseKey(n) === caseKey(words.FALSE) ||
    (cell !== null && colToIndex(cell[1]) < MAX_COLS && inGrid(cell[2], MAX_ROWS)) ||
    r1c1Start(n)
  return quote ? `'${n.replaceAll("'", "''")}'` : n
})

/**
 * The key a name compares by: sheets, tables and defined names are the same name in any case (`Sales`, `SALES`), by
 * Unicode's upper case in every locale (`straße` is `STRASSE`).
 */
export const nameKey = /* @__PURE__ */ strict(function nameKey(name: string): string {
  return caseKey(text(name, 'name'))
})

/** Whether two names of sheets, tables or defined names are the same name (`nameKey`). */
export const sameName = /* @__PURE__ */ strict(function sameName(a: string, b: string): boolean {
  return nameKey(a) === nameKey(b)
})
