/**
 * Edits of cells by their fields (`Session.setCells`): an edit's fields go over the cell's input, as Excel keeps a cell's
 * look when a value is typed over it. The content is one whole, a value or a formula with its array area: a new value
 * takes the formula's place and a new formula the value's; an edit of null clears the content and the hyperlink, as Delete.
 */
import type { CellEdit, CellFields, CellInput, CellReplace } from '../contract.js'
import { sheetIndexOf, type EngineState } from '../engine.js'
import { CELL_FIELDS, checkCell, checkEditList, InputError, object, onlyKeys } from '../input.js'
import { inputAt } from './book.js'

/** The fields of a cell's content: a formula's value held while a circular reference stops it goes with the formula. */
export const CONTENT = ['value', 'formula', 'array', 'held'] as const

/**
 * A cell's input with an edit's fields over it; null when nothing is left of the cell. An edit of null clears the content
 * and the hyperlink, as Excel's Delete key and ClearContents do (the set `lines/clear`).
 */
export function withFields(input: CellInput | undefined, fields: CellFields | null): CellInput | null {
  // a new value or formula, or none, replaces the content whole; a formula removed takes its area along
  const whole = fields === null || fields.value !== undefined || typeof fields.formula === 'string'
  const out: Record<string, unknown> = {}
  // the input's fields that stay, in their order, copied rather than deleted: a deleted field slows every later read of the object
  if (input)
    for (const k in input) {
      if ((whole && (CONTENT as readonly string[]).includes(k)) || (fields === null && k === 'link') || (fields?.formula === null && (k === 'array' || k === 'held')) || fields?.[k as keyof CellFields] === null) continue
      out[k] = input[k as keyof CellInput]
    }
  if (fields)
    for (const k in fields) {
      const v = fields[k as keyof CellFields]
      if (v !== null && v !== undefined) out[k] = v
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
    return [{ sheet: e.sheet, cell: e.cell, input: withFields(inputAt(s, sheetIndexOf(s, e.sheet)!, e.cell), e.input) }]
  }
  const now = new Map<string, CellInput | null>()
  return edits.map((e) => {
    const i = sheetIndexOf(s, e.sheet)!
    const k = `${i}!${e.cell}`
    const was = now.has(k) ? now.get(k) : inputAt(s, i, e.cell)
    const input = withFields(was ?? undefined, e.input)
    now.set(k, input)
    return { sheet: e.sheet, cell: e.cell, input }
  })
}
