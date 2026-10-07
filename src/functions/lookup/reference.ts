/**
 * References and positions. CHOOSE, COLUMNS and ROWS are ported from formula.js 4.6.1
 * (src/lookup-reference.js), Copyright (c) 2014 Sutoiku, Inc., MIT License (THIRD_PARTY.md); the rest is
 * this engine's own (formula.js gets values, not references).
 */
import { colToIndex, indexToCol, MAX_COLS, MAX_ROWS } from '../../model.js'
import type { Arg, Cx, FnDef, Grid } from '../../args.js'
import { arr } from '../../arrays/value.js'
import { Areas, E, isErr, Rng, type V } from '../../values.js'

/** A name that reads as an A1 address within the grid, an R1C1 reference or a boolean. */
function referenceLike(s: string): boolean {
  const a1 = /^([A-Za-z]{1,3})(\d+)$/.exec(s)
  if (a1) return colToIndex(a1[1]) < MAX_COLS && Number(a1[2]) >= 1 && Number(a1[2]) <= MAX_ROWS
  // R1C1: a name starting with R, C or RC followed by a digit, or just R, C, RC
  return /^(RC?|C)(\d|$)/i.test(s) || /^(TRUE|FALSE)$/i.test(s)
}

/** Excel quotes a sheet name unless it is empty or a plain name that does not read as a reference. */
const sheetPrefix = (s: string) => (s === '' || (/^[\p{L}_][\p{L}\p{N}_.]*$/u.test(s) && !referenceLike(s)) ? s : `'${s.replace(/'/g, "''")}'`)

function address(r: number, c: number, abs = 1, a1 = true, sheet?: string): V {
  if (r < 1 || r > MAX_ROWS || c < 1 || c > MAX_COLS || abs < 1 || abs > 4) return E.VALUE
  const absRow = abs <= 2
  const absCol = abs % 2 === 1
  const ref = a1 ? `${absCol ? '$' : ''}${indexToCol(c - 1)}${absRow ? '$' : ''}${r}` : `R${absRow ? r : `[${r}]`}C${absCol ? c : `[${c}]`}`
  return sheet === undefined ? ref : `${sheetPrefix(sheet)}!${ref}`
}

/** Of the formula's cell without an argument, else the reference's rows (columns): a vertical (horizontal) array of several. */
const position = (row: boolean): FnDef => ({
  args: 'R?',
  cell: true,
  array: true,
  fn: (_s: number, r: number, c: number, ref?: V | Rng) => {
    if (ref === undefined) return (row ? r : c) + 1
    if (!(ref instanceof Rng)) return isErr(ref) ? ref : E.VALUE
    const from = (row ? ref.r1 : ref.c1) + 1
    const n = row ? ref.r2 - ref.r1 + 1 : ref.c2 - ref.c1 + 1
    if (n === 1) return from
    const v = Array.from({ length: n }, (_, i) => from + i)
    return row ? arr(n, 1, v) : arr(1, n, v)
  },
})

export const functions: Record<string, FnDef> = {
  ADDRESS: { args: 'w w w? b? s?', fn: address },
  // the areas of a union; a reference across sheets, a value or an array is #VALUE!
  AREAS: { args: 'R', areas: true, fn: (x: Arg) => (isErr(x) ? x : x instanceof Rng ? 1 : x instanceof Areas && !x.cube ? x.list.length : E.VALUE) },
  // the chosen value as it is, a reference too, one cell too (SUMIF(CHOOSE(1,A1:A9,B1:B9),">2"))
  CHOOSE: { args: 'i P', rep: 1, fn: (i: number, values: V[]) => (i < 1 || i > values.length ? E.VALUE : values[i - 1]) },
  COLUMN: position(false),
  // the referenced size: a whole column has 1,048,576 rows
  COLUMNS: { args: 'G', fn: (g: Grid) => g.cols },
  // the value of a cell is its friendly name as given, else the link as text
  HYPERLINK: { args: 'v v?', cx: true, fn: (cx: Cx, link: V, name?: V) => (name === undefined ? cx.rules.toText(link) : name) },
  ROW: position(true),
  ROWS: { args: 'G', fn: (g: Grid) => g.rows },
}
