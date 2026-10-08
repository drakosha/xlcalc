/** The values a structure operation changed in a session: compared with the old value of the cell that moved there. */
import type { CellRef } from '../contract.js'
import { key, type EngineState } from '../engine.js'
import { same } from '../evaluate.js'
import { spend } from '../exec/budget.js'
import { formatAddress, parseAddress } from '../model.js'
import type { Looks } from '../shown.js'
import { shownKey } from './looks.js'
import { describeAt } from '../spill/place.js'
import { sheetKey } from '../syntax.js'
import type { V } from '../values.js'

/**
 * Cells of the new state whose value differs from the one their old cell had (`mapCell` carries each old value
 * to its new place; a cell nothing came to was empty), and formulas whose spill differs from their old cell's
 * (`describeAt`, which stands for a large array's unstored cells), as `s:r:c` keys of the new state in order;
 * with the looks of both states, also the cells that show something else than their old cell showed.
 */
export function movedChanges(old: EngineState, next: EngineState, mapCell: (ref: CellRef) => CellRef | null, looks?: [Looks, Looks]): string[] {
  const carried = new Map<string, V>()
  const shown = new Map<string, string>()
  old.ctx.sheets.forEach((sh, s) =>
    sh.cols.forEach((col, c) =>
      col?.forEach((v, r) => {
        spend(1)
        const to = v === null ? null : mapCell({ sheet: old.sheetNames[s], cell: formatAddress(r, c) })
        if (!to) return
        const p = parseAddress(to.cell)
        const k = `${next.sheetIdx.get(sheetKey(to.sheet))}:${p.row}:${p.col}`
        carried.set(k, v)
        if (looks) shown.set(k, shownKey(looks[0], old, s, r, c))
      }),
    ),
  )
  const out: number[][] = []
  next.ctx.sheets.forEach((sh, s) =>
    sh.cols.forEach((col, c) =>
      col?.forEach((v, r) => {
        spend(1)
        const k = `${s}:${r}:${c}`
        if (!same(carried.get(k) ?? null, v) || (looks && v !== null && shown.get(k) !== shownKey(looks[1], next, s, r, c))) out.push([s, r, c])
        carried.delete(k)
      }),
    ),
  )
  // a value carried to a cell that is empty now
  for (const k of carried.keys()) out.push(k.split(':').map(Number))
  const spills = new Map<string, string>()
  for (const x of old.spills.on.values()) {
    const to = mapCell({ sheet: old.sheetNames[x.s], cell: formatAddress(x.r1, x.c1) })
    if (!to) continue
    const p = parseAddress(to.cell)
    spills.set(key(next.sheetIdx.get(sheetKey(to.sheet))!, p.row, p.col), describeAt(old, key(x.s, x.r1, x.c1)))
  }
  for (const x of next.spills.on.values()) spills.set(key(x.s, x.r1, x.c1), spills.get(key(x.s, x.r1, x.c1)) ?? '')
  for (const [k, was] of spills) if (was !== describeAt(next, k) && !out.some((o) => o.join(':') === k)) out.push(k.split(':').map(Number))
  return out.sort((a, b) => a[0] - b[0] || a[1] - b[1] || a[2] - b[2]).map((k) => k.join(':'))
}
