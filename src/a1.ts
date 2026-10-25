/**
 * A1 addresses without the engine (the `./a1` subpath, the same functions in the root): cells, ranges, the keys of layout
 * lines and their maps patched, a sheet's outline groups and Excel's commands on them, rectangles, a sheet's name in a
 * reference as Excel writes it and the rule names compare by. Rows and columns are numbered from 1, as A1 numbers them;
 * the parts reused are the engine's own (`model.ts`, `structure/area.ts`, `layout/spans.ts`, `layout/outline.ts`,
 * `structure/outline-command.ts`, `collate.ts`, `input/basic.ts`), which compute nothing. Each function refuses options it does not take (`strict-args.ts`).
 */
import type { CellPosition, CellRect, LinePatch, LineSpan, OutlineEditOp, OutlineEditResult, OutlineGroup, RangeRect, SheetInput, SheetOutline } from './contract.js'
import { caseKey } from './collate.js'
import { InputError, what } from './input-error.js'
import { checkString, isObject } from './input/basic.js'
import { groupsOf, type OutlineEntry } from './layout/outline.js'
import { lines, parseSpan, patchLines as patchSpans, spanKey } from './layout/spans.js'
import type { Locale } from './locale/types.js'
import { OUTLINE_MESSAGES } from './locale/outline-messages.js'
import { colToIndex, formatAddress, indexToCol, MAX_COLS, MAX_ROWS } from './model.js'
import { contains, encode, grown, meets, parseRange as rangeArg, rectOf, span, type Area, type Axis, type Rect } from './structure/area.js'
import { applyOutline, OUTLINE_FIELDS, type OutlineAxes, type OutlineOp } from './structure/outline-command.js'
import { PLAIN_NAME, quoteSheet } from './sheet-pattern.js'
import { strict } from './strict-args.js'

export { MAX_COLS, MAX_ROWS }

/** A whole number from 1 to `max`, else `InputError` naming `path`: `checkNumber`'s rule written out, which costs the `./a1` bundle 80 bytes fewer than the general check. */
function line(n: unknown, max: number, path: string): number {
  if (typeof n !== 'number' || !Number.isInteger(n) || n < 1 || n > max) throw new InputError(path, `expected an integer from 1 to ${max}, got ${what(n)}`)
  return n
}

/** The letters of column `n` (1 is `A`, 16,384 is `XFD`). */
export const colLetters = /* @__PURE__ */ strict(function colLetters(n: number): string {
  return indexToCol(line(n, MAX_COLS, 'n') - 1)
})

/** The number of a column from its letters, in any case (`B` is 2), or null when they are not a column of the sheet (`XFE`). */
export const colNumber = /* @__PURE__ */ strict(function colNumber(letters: string): number | null {
  if (!/^[A-Za-z]{1,3}$/.test(checkString(letters, 'letters'))) return null
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

/** The engine's rectangle (rows and columns from 0) as a `CellRect` (from 1). */
const cellRect = (x: Rect): CellRect => ({ top: x.r1 + 1, left: x.c1 + 1, bottom: x.r2 + 1, right: x.c2 + 1 })

/**
 * A range's rectangle and kind: `B2:D9` and `B2` are cells, `B:D` entire columns, `2:9` entire rows, with `$` and
 * letters in any case; corners given in reverse order are swapped, as Excel reorders such a typed reference. Null for
 * anything else, or for a range that extends beyond the sheet.
 */
export const parseRange = /* @__PURE__ */ strict(function parseRange(range: string): RangeRect | null {
  // letters of other scripts stay out: some come to Latin ones in upper case (`ß` is `SS`)
  if (!/^[A-Za-z0-9$:]+$/.test(checkString(range, 'range'))) return null
  const t = range.replace(DOLLAR, '$1$2').toUpperCase()
  const x = rangeArg(t)
  if (!x) return null
  const kind = COLUMNS.test(t) ? 'columns' : ROWS.test(t) ? 'rows' : 'cells'
  return { kind, ...cellRect(x) }
})

/** A cell's row and column (`B12`, `$B$12`, `b12`), or null when the text is not a cell of the sheet (`ZZZ1`, `A0`, `B2:C3`). */
export const parseCell = /* @__PURE__ */ strict(function parseCell(cell: string): CellPosition | null {
  const r = checkString(cell, 'cell').includes(':') ? null : parseRange(cell)
  return r?.kind === 'cells' ? { row: r.top, col: r.left } : null
})

/** A rectangle checked against the sheet, corners in order, as the engine's area of a reference without `$`. */
function areaOf(r: CellRect, path = 'rect'): Area {
  // `?.`: what is no rectangle is refused by its first bound
  const [top, bottom] = [line(r?.top, MAX_ROWS, `${path}.top`), line(r?.bottom, MAX_ROWS, `${path}.bottom`)]
  const [left, right] = [line(r?.left, MAX_COLS, `${path}.left`), line(r?.right, MAX_COLS, `${path}.right`)]
  const b = (i: number) => ({ i: i - 1, abs: false })
  return { kind: top === bottom && left === right ? 'cell' : 'range', r: span(b(top), b(bottom)), c: span(b(left), b(right)) }
}

/** A rectangle's address as Excel writes it: `B2` for a single cell, `B2:D9`, `B:D` when it spans every row, `2:9` when it spans every column. */
export const rangeAddress = /* @__PURE__ */ strict(function rangeAddress(rect: CellRect): string {
  return encode(areaOf(rect))
})

/** A rectangle checked against the sheet, corners in order (rows and columns from 0), as every helper here takes one. */
const rect = (x: CellRect, path?: string): Rect => rectOf(areaOf(x, path))

/** Whether two rectangles share a cell. */
export const rectsMeet = /* @__PURE__ */ strict(function rectsMeet(a: CellRect, b: CellRect): boolean {
  return meets(rect(a, 'a'), rect(b, 'b'))
})

/** Whether `inner` lies within `outer`. */
export const rectContains = /* @__PURE__ */ strict(function rectContains(outer: CellRect, inner: CellRect): boolean {
  return contains(rect(outer, 'outer'), rect(inner, 'inner'))
})

/**
 * The rectangle a merge covers (`mergeCells`, Merge Cells): `rect` extended to include each merged area of `merges`
 * that it overlaps, repeated for each area it then overlaps, until it overlaps no other.
 */
export const mergedArea = /* @__PURE__ */ strict(function mergedArea(r: CellRect, merges: readonly CellRect[]): CellRect {
  const x = rect(r)
  if (!Array.isArray(merges)) throw new InputError('merges', `expected an array of rectangles, got ${what(merges)}`)
  return cellRect(grown(x, merges.map((m, i) => rect(m, `merges[${i}]`))))
})

/** The engine's axis of layout lines. */
const axisOf = (axis: 'columns' | 'rows') => {
  if (axis !== 'columns' && axis !== 'rows') throw new InputError('axis', `expected "columns" or "rows", got ${what(axis)}`)
  return axis === 'columns' ? 'c' : 'r'
}

/**
 * The lines that a key of `SheetInput.columns` or `rows` refers to (`C`, `B:D`; `5`, `5:9`), or null for a key a
 * workbook does not accept: lowercase, a span of a single line (`C:C`) or in reverse order, or beyond the sheet.
 */
export const parseLines = /* @__PURE__ */ strict(function parseLines(key: string, axis: 'columns' | 'rows'): LineSpan | null {
  const s = parseSpan(checkString(key, 'key'), axisOf(axis))
  return s && { first: s[0] + 1, last: s[1] + 1 }
})

/** The key for lines `first` to `last` (`C`, `B:D`; `5`, `5:9`), in ascending order regardless of which bound is given first. */
export const linesKey = /* @__PURE__ */ strict(function linesKey(first: number, last: number, axis: 'columns' | 'rows'): string {
  const ax = axisOf(axis)
  const max = ax === 'c' ? MAX_COLS : MAX_ROWS
  const [a, b] = [line(first, max, 'first'), line(last, max, 'last')]
  return spanKey(Math.min(a, b) - 1, Math.max(a, b) - 1, ax)
})

/**
 * A map of `SheetInput.columns` or `rows` with `patch` applied to the lines `span` (`C`, `B:D`; `5`, `5:9`; the axis
 * follows from the key's shape), as `setLayout` applies it: a value sets a field, null removes it, and an omitted field
 * is left unchanged; spans that the patch cuts through are split, and lines left without fields are dropped. Returns a
 * new map; the given one is not modified. The map is checked as a workbook's would be: an object for each span of the
 * axis, with no two spans overlapping.
 */
export const patchLines = /* @__PURE__ */ strict(function patchLines<T extends object>(map: Record<string, T> | undefined, span: string, patch: LinePatch<T>): Record<string, T> {
  const axis: Axis = /^[A-Z]/.test(checkString(span, 'span')) ? 'c' : 'r'
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

/** A sheet or its layout, its maps of rows and columns checked as `checkMap` checks them. */
function checkAxes(sheet: OutlineAxes) {
  if (!isObject(sheet)) throw new InputError('sheet', `expected a sheet or its layout, got ${what(sheet)}`)
  checkMap(sheet.rows, 'r', 'sheet.rows', 'a span of rows ("5", "5:9")')
  checkMap(sheet.columns, 'c', 'sheet.columns', 'a span of columns ("C", "B:D")')
}

/**
 * A sheet's groups of rows and columns as Excel draws them (`SheetOutline`), from its lines' outline levels, hidden
 * state and collapsed markers and its summary settings: a `SheetInput`, or a `SheetLayout` as `getLayout()` returns it.
 * Each axis's groups are ordered by first line, the outer before the inner when two start together; a group is
 * collapsed when Excel would keep the marker on its summary line (`RowInput.collapsed`), whatever markers the sheet
 * holds.
 */
export const outlineGroups = /* @__PURE__ */ strict(function outlineGroups(sheet: Pick<SheetInput, 'rows' | 'columns' | 'outline'>): SheetOutline {
  checkAxes(sheet)
  const o = sheet.outline
  const groups = (map: Record<string, OutlineEntry> | undefined, axis: Axis, after: boolean): OutlineGroup[] =>
    groupsOf(map, axis, after).map((g) => ({ first: g.lo + 1, last: g.hi + 1, level: g.level, summary: g.summary === null ? null : g.summary + 1, collapsed: g.collapsed }))
  return { rows: groups(sheet.rows, 'r', o?.summaryBelow !== false), columns: groups(sheet.columns, 'c', o?.summaryRight !== false) }
})

/**
 * Excel's outline commands on a standalone sheet, as `editStructure` performs them in a workbook (`OutlineEditOp`:
 * Group, Ungroup, Show and Hide Detail, the level buttons): the sheet's rows and columns after the command, taking its
 * summary settings into account, or why Excel rejects the command, in English
 * (`{ ok: false, code: 'detail', message: 'the detail of row 11 is already shown' }`). The sheet is checked as
 * `outlineGroups` checks it, and the operation as `editStructure` checks it.
 */
export const outlineEdit = /* @__PURE__ */ strict(function outlineEdit(sheet: Pick<SheetInput, 'rows' | 'columns' | 'outline'>, op: OutlineEditOp): OutlineEditResult {
  checkAxes(sheet)
  if (!isObject(op)) throw new InputError('op', `expected an outline operation, got ${what(op)}`)
  const fields = Object.hasOwn(OUTLINE_FIELDS, op.op as string) ? OUTLINE_FIELDS[op.op as OutlineOp['op']].filter((k) => k !== 'sheet') : null
  if (!fields) throw new InputError('op.op', `expected one of ${Object.keys(OUTLINE_FIELDS).join(', ')}, got ${what(op.op)}`)
  for (const k of Object.keys(op)) if (k !== 'op' && !fields.includes(k)) throw new InputError('op', `unknown field ${JSON.stringify(k)} (expected op, ${fields.join(', ')})`)
  const r = applyOutline(sheet, op)
  if (!r.ok) return { ok: false, code: r.code, message: OUTLINE_MESSAGES[r.message.key](r.message) }
  const { rows, columns } = r.sheet
  return { ok: true, ...(rows && Object.keys(rows).length ? { rows } : {}), ...(columns && Object.keys(columns).length ? { columns } : {}) }
})

/** An R1C1 reference at the start of a name: `R` and its number, then `C` and its number; or `C` and its number. */
const R1C1 = /^(?:([Rr])(\d*)(?:[Cc](\d*))?|[Cc](\d*))/
const CELL = /^([A-Za-z]{1,3})(\d+)$/

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

/** The booleans of `locale`, en-US's without one; its object alone, as a tag needs the list of locales this subpath does not carry. */
function booleans(locale: unknown): Locale['booleans'] {
  if (locale === undefined) return { TRUE: 'TRUE', FALSE: 'FALSE' }
  const b = (locale as Partial<Locale> | null)?.booleans
  if (typeof locale === 'object' && typeof b?.TRUE === 'string' && typeof b.FALSE === 'string') return b
  const hint = typeof locale === 'string' ? ' (it takes no tag, so that the a1 subpath carries no list of locales)' : ''
  throw new InputError('locale', `expected a locale object such as ruRU, with its booleans, got ${what(locale)}${hint}`)
}

/**
 * A sheet's name in a reference as Excel writes it: quoted, with any apostrophe inside doubled, when it is not a plain
 * name (`'Q1 data'`, `'2024'`, `'O''Brien'`), when it is a cell address on the sheet (`'A1'`, `'XFD1048576'`; not
 * `XFE1`), when it begins with an R1C1 reference (`'R'`, `'R1C1'`, `'R2D2'`) or when it is one of the locale's booleans
 * (`'TRUE'`; `'ИСТИНА'` with ru-RU). `locale` provides the booleans, en-US's (the stored text, `Range.Formula`) without
 * one: a locale object (`ruRU`), not a tag, so that `xlcalc/a1` carries no list of locales; a tag throws an
 * `InputError`. The engine's stored formulas quote every sheet name: this is the text a user reads.
 */
export const quoteSheetIfNeeded = /* @__PURE__ */ strict(function quoteSheetIfNeeded(name: string, locale?: Pick<Locale, 'booleans'>): string {
  const n = checkString(name, 'name')
  const words = booleans(locale)
  const cell = CELL.exec(n)
  const quote =
    !PLAIN_NAME.test(n) ||
    caseKey(n) === caseKey(words.TRUE) ||
    caseKey(n) === caseKey(words.FALSE) ||
    (cell !== null && colToIndex(cell[1]) < MAX_COLS && inGrid(cell[2], MAX_ROWS)) ||
    r1c1Start(n)
  return quote ? quoteSheet(n) : n
})

/**
 * The key by which names are compared: sheet, table and defined names are the same name in any case (`Sales`,
 * `SALES`), by Unicode uppercase in every locale (`straße` is `STRASSE`).
 */
export const nameKey = /* @__PURE__ */ strict(function nameKey(name: string): string {
  return caseKey(checkString(name, 'name'))
})

/** Whether two names of sheets, tables or defined names are the same name (`nameKey`). */
export const sameName = /* @__PURE__ */ strict(function sameName(a: string, b: string): boolean {
  return nameKey(a) === nameKey(b)
})
