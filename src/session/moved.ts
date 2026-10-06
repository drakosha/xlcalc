/** The values a structure operation changed in a session: compared with the old value of the cell that moved there. */
import type { CellRef } from '../contract.js'
import type { EngineState } from '../engine.js'
import { same } from '../evaluate.js'
import { formatAddress, parseAddress } from '../model.js'
import { sheetKey } from '../syntax.js'
import type { V } from '../values.js'

/**
 * Cells of the new state whose value differs from the one their old cell had (`mapCell` carries each old value
 * to its new place; a cell nothing came to was empty), as `s:r:c` keys of the new state in order.
 */
export function movedChanges(old: EngineState, next: EngineState, mapCell: (ref: CellRef) => CellRef | null): string[] {
  const carried = new Map<string, V>()
  old.ctx.sheets.forEach((sh, s) =>
    sh.cols.forEach((col, c) =>
      col?.forEach((v, r) => {
        const to = v === null ? null : mapCell({ sheet: old.sheetNames[s], cell: formatAddress(r, c) })
        if (!to) return
        const p = parseAddress(to.cell)
        carried.set(`${next.sheetIdx.get(sheetKey(to.sheet))}:${p.row}:${p.col}`, v)
      }),
    ),
  )
  const out: number[][] = []
  next.ctx.sheets.forEach((sh, s) =>
    sh.cols.forEach((col, c) =>
      col?.forEach((v, r) => {
        const k = `${s}:${r}:${c}`
        if (!same(carried.get(k) ?? null, v)) out.push([s, r, c])
        carried.delete(k)
      }),
    ),
  )
  // a value carried to a cell that is empty now
  for (const k of carried.keys()) out.push(k.split(':').map(Number))
  return out.sort((a, b) => a[0] - b[0] || a[1] - b[1] || a[2] - b[2]).map((k) => k.join(':'))
}
