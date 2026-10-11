/**
 * Patches of looks (`Session.patchCells`) checked: a sheet of the book and a cell or a range of it, and the fields to
 * set: a style by its number of the book's table or by its fields (`StylePatch`), the number format, meta, a note, a
 * link. On whole rows and columns meta, notes and links can only be removed.
 */
import type { CellLink, CellNote, JsonValue, NumberFormat, StylePatch } from '../contract.js'
import { array, checkAddress, checkCell, checkFormat, checkSheet, checkString, checkStyle, checkStyleIndex, InputError, object, onlyKeys } from '../input.js'
import { MAX_COLS, MAX_ROWS } from '../model.js'
import { parseRange, type Rect } from '../structure/area.js'

/** What a patch does to a style: one of the table's set, fields changed over the style a cell shows, the style removed, or the look reset (Excel's ClearFormats). */
export type StyleChange = { set: number } | { merge: StylePatch } | { remove: true } | { reset: true }

/** What a patch does to a number format: one set, removed, or reset to General (Excel's ClearFormats). */
export type FormatChange = { set: NumberFormat } | { remove: true } | { reset: true }

/** What a patch does to the cells it covers; `content` clears their values and formulas (`Session.clearCells`). */
export interface Look {
  style?: StyleChange
  format?: FormatChange
  meta?: JsonValue | null
  note?: CellNote | null
  link?: CellLink | null
  content?: true
}

/** Cells of a sheet: a rectangle, or whole rows (`r`) or columns (`c`). */
export interface Target {
  i: number
  rect: Rect
  lines?: 'r' | 'c'
}

export interface CheckedPatch extends Target {
  look: Look
}

const FIELDS = ['style', 'numberFormat', 'meta', 'note', 'link'] as const

/** Whole rows, whole columns, or neither: a rectangle across the sheet's width is rows, as Excel takes it. */
export const linesOf = (x: Rect): 'r' | 'c' | undefined => (x.c1 === 0 && x.c2 === MAX_COLS - 1 ? 'r' : x.r1 === 0 && x.r2 === MAX_ROWS - 1 ? 'c' : undefined)

/** A range of a sheet (`B2`, `B2:D9`, `B:D`, `2:9`). */
export function checkRange(v: unknown, path: string): Rect {
  const x = parseRange(checkString(v, path))
  if (!x) throw new InputError(path, `expected a range such as "B2", "B2:D9", "B:D" or "2:9" within the sheet, got ${JSON.stringify(v)}`)
  return x
}

/** Patches of a book with `styles` styles, each with its target and what it does there. */
export function checkPatches(v: unknown, sheetIndex: (name: string) => number | undefined, styles: number, path = 'patches'): CheckedPatch[] {
  return array(v, path, 'an array of patches such as [{ sheet: "Sheet1", range: "B2:D9", set: { style: { font: { bold: true } } } }]').map((e, k) => {
    const p = `${path}[${k}]`
    const o = object(e, p, 'a patch such as { sheet: "Sheet1", range: "B2:D9", set: { numberFormat: "0.00" } }')
    onlyKeys(o, p, ['sheet', 'cell', 'range', 'set'])
    const i = checkSheet(o.sheet, sheetIndex, `${p}.sheet`)
    if ((o.cell === undefined) === (o.range === undefined)) throw new InputError(p, 'expected a cell or a range, one of them')
    const rect = o.cell !== undefined ? parseRange(checkAddress(o.cell, `${p}.cell`))! : checkRange(o.range, `${p}.range`)
    const set = object(o.set, `${p}.set`, 'the fields to change, such as { style: { font: { bold: true } }, numberFormat: "0.00", note: null }')
    onlyKeys(set, `${p}.set`, FIELDS)
    const lines = linesOf(rect)
    const look: Look = {}
    const style = set.style
    if (style === null) look.style = { remove: true }
    else if (typeof style === 'number') {
      checkStyleIndex(style, `${p}.set.style`, styles)
      look.style = { set: style }
    } else if (style !== undefined) look.style = { merge: checkStyle(style, `${p}.set.style`, true) as StylePatch }
    if (set.numberFormat === null) look.format = { remove: true }
    else if (set.numberFormat !== undefined) {
      checkFormat(set.numberFormat, `${p}.set.numberFormat`)
      look.format = { set: set.numberFormat as NumberFormat }
    }
    const given: Record<string, unknown> = {}
    for (const key of ['meta', 'note', 'link'] as const) {
      if (set[key] === undefined) continue
      if (set[key] !== null && lines) throw new InputError(`${p}.set.${key}`, `whole ${lines === 'r' ? 'rows' : 'columns'} take no ${key}: null removes it from their cells`)
      look[key] = set[key] as never
      if (set[key] !== null) given[key] = set[key]
    }
    checkCell(given, `${p}.set`, styles)
    return { i, rect, ...(lines ? { lines } : {}), look }
  })
}
