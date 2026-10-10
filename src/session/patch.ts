/**
 * Patches of cells (`Session.patchCells`): a cell's style, format, meta, note and link changed without its value or
 * formula, which the patch does not carry. Nothing is recomputed: a value never depends on these.
 */
import type { CellInput, CellPatch } from '../contract.js'
import { setInput, type EngineState } from '../engine.js'
import { array, checkAddress, checkCell, checkSheet, object, onlyKeys } from '../input.js'
import { parseAddress } from '../model.js'

const FIELDS = ['style', 'format', 'meta', 'note', 'link'] as const

/** Patches of a book with `styles` styles: a sheet of the book, a cell, and the fields to set as a cell's (null removing one). */
export function checkPatches(v: unknown, sheetIndex: (name: string) => number | undefined, styles: number, path = 'patches'): CellPatch[] {
  return array(v, path, 'an array of patches such as [{ sheet: "Sheet1", cell: "A1", set: { style: 0 } }]').map((e, i) => {
    const p = `${path}[${i}]`
    const o = object(e, p, 'a patch such as { sheet: "Sheet1", cell: "A1", set: { style: 0 } }')
    onlyKeys(o, p, ['sheet', 'cell', 'set'])
    checkSheet(o.sheet, sheetIndex, `${p}.sheet`)
    checkAddress(o.cell, `${p}.cell`)
    const set = object(o.set, `${p}.set`, 'the fields to change, such as { style: 0, format: "0.00", note: null }')
    onlyKeys(set, `${p}.set`, FIELDS)
    checkCell(Object.fromEntries(Object.entries(set).filter(([, x]) => x !== null)), `${p}.set`, styles)
    return o as unknown as CellPatch
  })
}

/** A cell's input with a patch applied; null when nothing is left in it. */
export function patched(input: CellInput | undefined, set: CellPatch['set']): CellInput | null {
  const out: Record<string, unknown> = { ...input }
  for (const k of FIELDS) {
    const x = set[k]
    if (x === null) delete out[k]
    else if (x !== undefined) out[k] = x
  }
  return Object.values(out).some((x) => x !== undefined) ? (out as CellInput) : null
}

/** The patches (checked) applied to the book's input, each on what the ones before it left. */
export function applyPatches(s: EngineState, patches: CellPatch[], sheetIndex: (name: string) => number) {
  const next = new Map<string, [number, string, CellInput | null]>()
  for (const p of patches) {
    const i = sheetIndex(p.sheet)
    const k = `${i}:${p.cell}`
    const was = next.has(k) ? next.get(k)![2] : s.input.sheets[i].cells[p.cell]
    next.set(k, [i, p.cell, patched(was ?? undefined, p.set)])
  }
  for (const [i, addr, input] of next.values()) {
    const { row, col } = parseAddress(addr)
    setInput(s, i, addr, row, col, input)
  }
}
