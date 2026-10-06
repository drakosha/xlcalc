/**
 * References and positions. CHOOSE, COLUMNS and ROWS are ported from formula.js 4.6.1
 * (src/lookup-reference.js), Copyright (c) 2014 Sutoiku, Inc., MIT License (THIRD_PARTY.md); the rest is
 * this engine's own (formula.js gets values, not references).
 */
import type { FnDef, Grid } from '../../args.js'
import { indexToCol, MAX_COLS, MAX_ROWS } from '../../model.js'
import { E, isErr, Rng, type V } from '../../values.js'

/** Excel quotes a sheet name that is not a plain name or that reads as a cell address. */
const sheetPrefix = (s: string) => (/^[\p{L}_][\p{L}\p{N}_.]*$/u.test(s) && !/^[A-Za-z]{1,3}\d+$/.test(s) && !/^R\d*C\d*$/i.test(s) ? s : `'${s.replace(/'/g, "''")}'`)

function address(r: number, c: number, abs = 1, a1 = true, sheet?: string): V {
  if (r < 1 || r > MAX_ROWS || c < 1 || c > MAX_COLS || abs < 1 || abs > 4) return E.VALUE
  const absRow = abs <= 2
  const absCol = abs % 2 === 1
  const ref = a1 ? `${absCol ? '$' : ''}${indexToCol(c - 1)}${absRow ? '$' : ''}${r}` : `R${absRow ? r : `[${r}]`}C${absCol ? c : `[${c}]`}`
  return sheet === undefined ? ref : `${sheetPrefix(sheet)}!${ref}`
}

/** Of the formula's cell without an argument, else the reference's first row (column), as Excel shows it in the cell. */
const position = (row: boolean): FnDef => ({
  args: 'R?',
  cell: true,
  fn: (_s: number, r: number, c: number, ref?: V | Rng) => {
    if (ref === undefined) return (row ? r : c) + 1
    if (ref instanceof Rng) return (row ? ref.r1 : ref.c1) + 1
    return isErr(ref) ? ref : E.VALUE
  },
})

export const functions: Record<string, FnDef> = {
  ADDRESS: { args: 'i i i? b? s?', fn: address },
  CHOOSE: { args: 'i a', rep: 1, fn: (i: number, values: V[]) => (i < 1 || i > values.length ? E.VALUE : values[i - 1]) },
  COLUMN: position(false),
  // the referenced size: a whole column has 1,048,576 rows
  COLUMNS: { args: 'G', fn: (g: Grid) => g.cols },
  // the value of a cell is its friendly name, else the link
  HYPERLINK: { args: 'v v?', fn: (link: V, name?: V) => (name === undefined ? link : name) },
  ROW: position(true),
  ROWS: { args: 'G', fn: (g: Grid) => g.rows },
}
