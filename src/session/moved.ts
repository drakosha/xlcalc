/** The values a structure operation changed in a session, and the formulas and looks: compared with those of the cell that moved there. */
import type { CellInput, CellRef } from '../contract.js'
import { key, type EngineState } from '../engine.js'
import { same } from '../evaluate.js'
import { spend } from '../exec/budget.js'
import { formatAddress, MAX_COLS, parseAddress } from '../model.js'
import { fieldsBetween, formulaText, type FieldChange } from './answer.js'
import type { Looks } from '../shown.js'
import { shownKey } from './looks.js'
import { describeAt } from '../spill/place.js'
import { sheetKey } from '../syntax.js'
import type { V } from '../values.js'

/**
 * Cells of the new state whose value differs from the one their old cell had (`mapCell` carries each old value
 * to its new place; a cell nothing came to was empty), and formulas whose spill differs from their old cell's
 * (`describeAt`, which stands for a large array's unstored cells), as `s:r:c` keys of the new state in order;
 * with the looks of both states, also the cells that show something else than their old cell showed. With `fields`,
 * also the cells whose formula (as each state shows it), own style or own number format differ from their old
 * cell's, by key, found in the same pass: a cell carried as it was costs nothing more.
 */
export function movedChanges(old: EngineState, next: EngineState, mapCell: (ref: CellRef) => CellRef | null, looks?: [Looks, Looks], fields?: Map<string, FieldChange>): string[] {
  const carried = new Map<string, V>()
  const shown = new Map<string, string>()
  const textOf = (st: EngineState, i: number, r: number, c: number, input: CellInput | undefined) => (input?.formula === undefined ? undefined : formulaText(st, i, r, c, input))
  /** The fields of a cell of the new state against those of the old cell that came to it (none: an empty one). */
  const compare = (k: string, j: number, r: number, c: number, now: CellInput | undefined, from?: [number, number, number, CellInput | undefined]) => {
    const was = from?.[3]
    // the same input without a formula: nothing of it can differ (a formula's text may, with new names)
    if (was === now && was?.formula === undefined) return
    const f = fieldsBetween(was, from && textOf(old, from[0], from[1], from[2], was), now, textOf(next, j, r, c, now))
    if (f) fields!.set(k, f)
  }
  /** An old cell carried to its new place: its key there, and with `fields` its fields compared. */
  const carry = (s: number, r: number, c: number): string | null => {
    const addr = formatAddress(r, c)
    const to = mapCell({ sheet: old.sheetNames[s], cell: addr })
    if (!to) return null
    const j = next.sheetIdx.get(sheetKey(to.sheet))!
    const p = parseAddress(to.cell)
    const k = `${j}:${p.row}:${p.col}`
    if (fields) compare(k, j, p.row, p.col, next.input.sheets[j].cells[to.cell], [s, r, c, old.input.sheets[s].cells[addr]])
    return k
  }
  old.ctx.sheets.forEach((sh, s) =>
    sh.cols.forEach((col, c) =>
      col?.forEach((v, r) => {
        spend(1)
        const k = v === null ? null : carry(s, r, c)
        if (k === null) return
        carried.set(k, v)
        if (looks) shown.set(k, shownKey(looks[0], old, s, r, c))
      }),
    ),
  )
  // the cells of a style, a format, a note or a link alone, which hold no value
  const reached = new Set<string>()
  if (fields)
    for (const [s, cells] of old.bare) for (const n of cells) {
      const k = carry(s, Math.floor(n / MAX_COLS), n % MAX_COLS)
      if (k !== null) reached.add(k)
    }
  const out: number[][] = []
  next.ctx.sheets.forEach((sh, s) =>
    sh.cols.forEach((col, c) =>
      col?.forEach((v, r) => {
        spend(1)
        const k = `${s}:${r}:${c}`
        if (fields && !carried.has(k) && !reached.has(k)) compare(k, s, r, c, next.input.sheets[s].cells[formatAddress(r, c)])
        if (!same(carried.get(k) ?? null, v) || (looks && v !== null && shown.get(k) !== shownKey(looks[1], next, s, r, c))) out.push([s, r, c])
        carried.delete(k)
      }),
    ),
  )
  if (fields)
    for (const [s, cells] of next.bare) for (const n of cells) {
      const [r, c] = [Math.floor(n / MAX_COLS), n % MAX_COLS]
      const k = `${s}:${r}:${c}`
      if (!carried.has(k) && !reached.has(k)) compare(k, s, r, c, next.input.sheets[s].cells[formatAddress(r, c)])
    }
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
