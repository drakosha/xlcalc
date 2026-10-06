/** What a cell of a result shows, for the spill tests: a large array's unstored cells are read from its `spillFill`. */
import type { Val } from '../bench/coverage/formulajs.js'
import type { ComputeResult } from '../src/index.js'
import { parseAddress } from '../src/model.js'
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
