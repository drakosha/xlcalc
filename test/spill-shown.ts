/** What a cell of a result shows, for the spill tests: a large array's unstored cells are read from its `spillFill`. */
import type { Val } from '../bench/coverage/formulajs.js'
import type { ComputeResult } from '../src/index.js'
import { formatAddress, parseAddress } from '../src/model.js'
import { toVal } from './excel-cases.js'

/** What a cell of the result shows: its value, a large array's unstored cells their fill, else blank. */
export function shownValue(r: Pick<ComputeResult, 'sheets'>, sheet: string, addr: string): Val | null {
  const cells = r.sheets.find((s) => s.name === sheet)!.cells
  const c = cells[addr]
  if (c) return c.value === null ? null : toVal(c.value)
  const { row, col } = parseAddress(addr)
  for (const x of Object.values(cells)) {
    if (x.spillFill === undefined) continue
    const [p, q] = x.spill!.split(':').map(parseAddress)
    if (row >= p.row && row <= q.row && col >= p.col && col <= q.col) return toVal(x.spillFill)
  }
  return null
}

/**
 * The cells of the rectangles `read` where the engine does not show Excel's values (`excel`, by `Sheet!A1`, none for a
 * blank cell), as `same` compares two values: `Sheet!A1: Excel …, engine …` lines.
 */
export function readDifferences(read: readonly { sheet: string; range: string }[], excel: Readonly<Record<string, Val>>, r: Pick<ComputeResult, 'sheets'>, same: (x: Val, engine: Val) => boolean): string[] {
  const out: string[] = []
  for (const g of read) {
    const [p, q = p] = g.range.split(':').map(parseAddress)
    for (let row = p.row; row <= q.row; row++)
      for (let col = p.col; col <= q.col; col++) {
        const addr = formatAddress(row, col)
        const x = excel[`${g.sheet}!${addr}`] ?? null
        const engine = shownValue(r, g.sheet, addr)
        const ok = x === null || engine === null ? x === engine : same(x, engine)
        if (!ok) out.push(`${g.sheet}!${addr}: Excel ${JSON.stringify(x)}, engine ${JSON.stringify(engine)}`)
      }
  }
  return out
}
