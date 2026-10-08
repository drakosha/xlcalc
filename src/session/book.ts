/** Changes of a session's book that recompute nothing: styles, formats and meta of cells, the layout of sheets. */
import type { BookInput, CellEdit, CellInput, LayoutEdit, SheetInput } from '../contract.js'
import type { EngineState } from '../engine.js'
import { InputError } from '../input.js'
import { parseSpan, patchLines } from '../layout/spans.js'
import { bookScalar } from '../model.js'
import { sheetKey } from '../syntax.js'

/** What a cell computes from: its formula, or its value (a date as its serial number). */
const content = (c: CellInput | null | undefined, epoch: number) => (c?.formula !== undefined ? `f${JSON.stringify([c.formula, c.array])}` : `v${JSON.stringify(bookScalar(c?.value, epoch))}`)

/**
 * The edits that change a cell's value or formula, the last one per cell; the others (a new style, format or
 * meta) are applied to the book's input here.
 */
export function contentEdits(s: EngineState, edits: CellEdit[]): CellEdit[] {
  const last = new Map<string, [CellEdit, Record<string, CellInput>]>()
  for (const e of edits) {
    const i = s.sheetIdx.get(sheetKey(e.sheet))!
    last.set(`${i}:${e.cell}`, [e, s.input.sheets[i].cells])
  }
  const out: CellEdit[] = []
  for (const [e, cells] of last.values()) {
    if (content(cells[e.cell], s.ctx.rules.epoch) !== content(e.input, s.ctx.rules.epoch)) out.push(e)
    else if (e.input) cells[e.cell] = e.input
    else delete cells[e.cell]
  }
  return out
}

/** Every style that cells, rows and columns use must be in a table of `count` styles. */
export function checkStylesUsed(book: BookInput, count: number) {
  const check = (style: number | undefined, where: string) => {
    if (style !== undefined && style >= count) throw new InputError('styles', `${where} uses style ${style}, the new table has ${count}`)
  }
  for (const s of book.sheets) {
    for (const [addr, c] of Object.entries(s.cells)) check(c.style, `${s.name}!${addr}`)
    for (const [k, c] of Object.entries(s.columns ?? {})) check(c.style, `columns ${k} of ${s.name}`)
    for (const [k, r] of Object.entries(s.rows ?? {})) check(r.style, `rows ${k} of ${s.name}`)
  }
}

/** Layout edits (checked) applied to the book: each changed sheet is a new object with the same cells. */
export function applyLayout(book: BookInput, edits: LayoutEdit[], sheetIndex: (name: string) => number) {
  const changed = new Map<number, SheetInput>()
  for (const e of edits) {
    const i = sheetIndex(e.sheet)
    const sheet = changed.get(i) ?? { ...book.sheets[i] }
    changed.set(i, sheet)
    if ('merges' in e) sheet.merges = [...e.merges]
    else if ('filtered' in e) sheet.filtered = e.filtered
    else if ('columns' in e) sheet.columns = patchLines(sheet.columns, 'c', ...parseSpan(e.columns, 'c')!, e.set)
    else sheet.rows = patchLines(sheet.rows, 'r', ...parseSpan(e.rows, 'r')!, e.set)
  }
  for (const [i, sheet] of changed) {
    for (const k of ['columns', 'rows', 'merges'] as const) if (sheet[k] && !Object.keys(sheet[k]).length) delete sheet[k]
    if (sheet.filtered === false) delete sheet.filtered
    book.sheets[i] = sheet
  }
}
