/**
 * Edits of cells by their fields (`Session.setCells`): an edit's fields go over the cell's input, as Excel keeps a cell's
 * look when a value is typed over it. The content is one whole, a value or a formula with its array area: a new value
 * takes the formula's place and a new formula the value's; an edit of null clears the content and the hyperlink, as Delete.
 */
import type { CellEdit, CellFields, CellInput, CellReplace } from '../contract.js'
import type { EngineState } from '../engine.js'
import { CELL_FIELDS, checkCell, checkEditList, InputError, object, onlyKeys } from '../input.js'
import { sheetKey } from '../syntax.js'
import { inputAt } from './book.js'

/** The fields of a cell's content: a formula's value held while a circular reference stops it goes with the formula. */
export const CONTENT = ['value', 'formula', 'array', 'held'] as const

/**
 * A cell's input with an edit's fields over it; null when nothing is left of the cell. An edit of null clears the content
 * and the hyperlink, as Excel's Delete key and ClearContents do (the set `lines/clear`).
 */
export function withFields(input: CellInput | undefined, fields: CellFields | null): CellInput | null {
  const out: Record<string, unknown> = { ...input }
  // a new value or formula, or none, replaces the content whole; a formula removed takes its area along
  if (fields === null || fields.value !== undefined || typeof fields.formula === 'string') for (const k of CONTENT) delete out[k]
  if (fields === null) delete out.link
  if (fields?.formula === null) {
    delete out.array
    delete out.held
  }
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
  return checkEditList(v, sheetIndex, path, 'the fields of a cell, or null to clear its content', (input, p) => {
    const f = object(input, p, 'the fields of a cell such as { value: 1 } or { formula: "A1*2", style: null }')
    // the fields given are checked as a cell's, those removed (null) left out; most edits remove none
    onlyKeys(f, p, CELL_FIELDS)
    let given = f
    for (const k in f)
      if (f[k] === null && k !== 'value') {
        given = Object.fromEntries(Object.entries(f).filter(([x, v]) => v !== null || x === 'value'))
        break
      }
    checkCell(given, p, styles)
    if (given.array !== undefined && typeof f.formula !== 'string') throw new InputError(`${p}.array`, 'an array area goes with its formula')
  })
}

/** Edits by fields as whole inputs, each over what the cell holds after the edits before it. */
export function resolveEdits(s: EngineState, edits: CellEdit[]): CellReplace[] {
  if (edits.length === 1) {
    const [e] = edits
    return [{ sheet: e.sheet, cell: e.cell, input: withFields(inputAt(s, s.sheetIdx.get(sheetKey(e.sheet))!, e.cell), e.input) }]
  }
  const now = new Map<string, CellInput | null>()
  return edits.map((e) => {
    const i = s.sheetIdx.get(sheetKey(e.sheet))!
    const k = `${i}!${e.cell}`
    const was = now.has(k) ? now.get(k) : inputAt(s, i, e.cell)
    const input = withFields(was ?? undefined, e.input)
    now.set(k, input)
    return { sheet: e.sheet, cell: e.cell, input }
  })
}
