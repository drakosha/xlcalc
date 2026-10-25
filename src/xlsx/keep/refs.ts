/**
 * References of the file's parts as a carrier holds them and as a file writes them: a carrier's text read back as its
 * areas (each with its sheet, or lost), an area written with the `$`s it had in the file, a sheet's name quoted where
 * Excel quotes it.
 */
import { colToIndex, indexToCol, MAX_COLS, MAX_ROWS } from '../../model.js'
import type { Rect } from '../../structure/area.js'
import { FormulaError, type Ast } from '../../syntax/ast.js'
import { plain } from '../../syntax/links.js'
import { parse } from '../../syntax/parse.js'
import { quoteSheet } from '../../syntax/print.js'

/** An area of a carrier's reference: its sheet and rectangle, or lost (`#REF!`), on its sheet or with the sheet. */
export type Area = { sheet: string; rect: Rect } | { lost: true; sheet?: string }

/** A carrier's text as its areas; null when it is no union of references. */
export function areasOf(text: string): Area[] | null {
  let ast: Ast
  try {
    ast = parse(text, { union: true })
  } catch (e) {
    if (e instanceof FormulaError) return null
    throw e
  }
  const out: Area[] = []
  for (const x of ast.k === 'union' ? ast.items : [ast]) {
    if (x.k === 'ref' && x.sheet !== null) out.push({ sheet: x.sheet, rect: { r1: x.ref.r1, r2: x.ref.r2, c1: x.ref.c1, c2: x.ref.c2 } })
    else if (x.k === 'err' && x.code === 'REF' && x.book === undefined) out.push(x.sheet === undefined ? { lost: true } : { lost: true, sheet: x.sheet })
    else return null
  }
  return out
}

/** The `$`s of an area as a file wrote it: before the first and the last cell's column and row. */
export interface Dollars {
  c1: boolean
  r1: boolean
  c2: boolean
  r2: boolean
}

export const ABSOLUTE: Dollars = { c1: true, r1: true, c2: true, r2: true }
const RELATIVE: Dollars = { c1: false, r1: false, c2: false, r2: false }

/** The `$`s of a reference's text without its sheet (`$B$2:C5`, `A:A`, `$1:$1`), as the parser keeps it (`RangeRef.text`). */
export function dollarsOf(ref: { text: string; kind: 'cell' | 'range' | 'cols' | 'rows' }): Dollars {
  const [a, b = a] = ref.text.split(':')
  const col = (t: string) => t.startsWith('$')
  const row = (t: string) => /\$\d/.test(t)
  if (ref.kind === 'cols') return { c1: col(a), r1: true, c2: col(b), r2: true }
  if (ref.kind === 'rows') return { c1: true, r1: col(a), c2: true, r2: col(b) }
  return { c1: col(a), r1: row(a), c2: col(b), r2: row(b) }
}

/** A rectangle as a file writes an area: a cell, a range, whole columns (`$C:$D`) or rows (`$3:$5`), with the `$`s given. */
export function areaText(x: Rect, d: Dollars = ABSOLUTE): string {
  const col = (c: number, abs: boolean) => `${abs ? '$' : ''}${indexToCol(c)}`
  const row = (r: number, abs: boolean) => `${abs ? '$' : ''}${r + 1}`
  if (x.r1 === 0 && x.r2 === MAX_ROWS - 1) return `${col(x.c1, d.c1)}:${col(x.c2, d.c2)}`
  if (x.c1 === 0 && x.c2 === MAX_COLS - 1) return `${row(x.r1, d.r1)}:${row(x.r2, d.r2)}`
  const first = col(x.c1, d.c1) + row(x.r1, d.r1)
  return x.r1 === x.r2 && x.c1 === x.c2 ? first : `${first}:${col(x.c2, d.c2)}${row(x.r2, d.r2)}`
}

/** An area without `$`s, as an `sqref` and a pivot table's place write it. */
export const plainArea = (x: Rect) => areaText(x, RELATIVE)

/** A sheet's name before `!`, quoted where Excel quotes it. */
export const sheetPrefix = (sheet: string) => (plain(sheet) ? sheet : quoteSheet(sheet))

/** The rectangles of an `sqref`: space-separated cells and ranges; null when one is none. */
export function sqrefRects(sqref: string): Rect[] | null {
  const out: Rect[] = []
  for (const p of sqref.trim().split(/\s+/).filter(Boolean)) {
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(p)
    if (!m) return null
    const [c1, r1] = [colToIndex(m[1].toUpperCase()), Number(m[2]) - 1]
    const [c2, r2] = m[3] ? [colToIndex(m[3].toUpperCase()), Number(m[4]) - 1] : [c1, r1]
    if (!(r1 >= 0 && r2 >= r1 && c2 >= c1 && r2 < MAX_ROWS && c2 < MAX_COLS)) return null
    out.push({ r1, r2, c1, c2 })
  }
  return out
}
