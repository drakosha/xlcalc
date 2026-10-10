/**
 * Edits of cells by their fields (`Session.setCells`): an edit's fields go over the cell's input, as Excel keeps a cell's
 * look when a value is typed over it. The content is one whole, a value or a formula with its array area: a new value
 * takes the formula's place and a new formula the value's; an edit of null clears the content alone, as Delete.
 */
import type { CellEdit, CellFields, CellInput, CellReplace } from '../contract.js'
import type { EngineState } from '../engine.js'
import { array, checkAddress, checkCell, checkSheet, InputError, object, onlyKeys } from '../input.js'
import { sheetKey } from '../syntax.js'

/** The fields of a cell's content. */
const CONTENT = ['value', 'formula', 'array'] as const

/** A cell's input with an edit's fields over it (null: the content cleared); null when nothing is left of the cell. */
export function withFields(input: CellInput | undefined, fields: CellFields | null): CellInput | null {
  const out: Record<string, unknown> = { ...input }
  // a new value or formula, or none, replaces the content whole; a formula removed takes its area along
  if (fields === null || fields.value !== undefined || typeof fields.formula === 'string') for (const k of CONTENT) delete out[k]
  if (fields?.formula === null) delete out.array
  if (fields)
    for (const k in fields) {
      const v = fields[k as keyof CellFields]
      if (v === null) delete out[k]
      else if (v !== undefined) out[k] = v
    }
  for (const k in out) if (out[k] !== undefined) return out as CellInput
  return null
}

/** Edits by fields of a book with `styles` styles: a sheet of the book, an address, the fields (null removing one) or null. */
export function checkFieldEdits(v: unknown, sheetIndex: (name: string) => number | undefined, styles: number, path = 'edits'): CellEdit[] {
  return array(v, path, 'an array of edits such as [{ sheet: "Sheet1", cell: "A1", input: { value: 1 } }]').map((e, i) => {
    const p = `${path}[${i}]`
    const o = object(e, p, 'an edit such as { sheet: "Sheet1", cell: "A1", input: { value: 1 } }')
    onlyKeys(o, p, ['sheet', 'cell', 'input'])
    checkSheet(o.sheet, sheetIndex, `${p}.sheet`)
    checkAddress(o.cell, `${p}.cell`)
    if (o.input === undefined) throw new InputError(`${p}.input`, 'expected the fields of a cell, or null to clear its content')
    if (o.input !== null) {
      const f = object(o.input, `${p}.input`, 'the fields of a cell such as { value: 1 } or { formula: "A1*2", style: null }')
      const given: Record<string, unknown> = {}
      for (const k in f) if (f[k] !== null || k === 'value') given[k] = f[k]
      checkCell(given, `${p}.input`, styles)
      if (given.array !== undefined && typeof f.formula !== 'string') throw new InputError(`${p}.input.array`, 'an array area goes with its formula')
    }
    return o as unknown as CellEdit
  })
}

/** Edits by fields as whole inputs, each over what the cell holds after the edits before it. */
export function resolveEdits(s: EngineState, edits: CellEdit[]): CellReplace[] {
  if (edits.length === 1) {
    const [e] = edits
    return [{ sheet: e.sheet, cell: e.cell, input: withFields(s.input.sheets[s.sheetIdx.get(sheetKey(e.sheet))!].cells[e.cell], e.input) }]
  }
  const now = new Map<string, CellInput | null>()
  return edits.map((e) => {
    const i = s.sheetIdx.get(sheetKey(e.sheet))!
    const k = `${i}!${e.cell}`
    const was = now.has(k) ? now.get(k) : s.input.sheets[i].cells[e.cell]
    const input = withFields(was ?? undefined, e.input)
    now.set(k, input)
    return { sheet: e.sheet, cell: e.cell, input }
  })
}
