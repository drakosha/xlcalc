/**
 * A spill case computed as Excel had it (`test/excel/spill/`, `test/excel/win/spill/`): the formulas a file stores as of
 * the old kind converted, the edits applied one by one in a session; and the cells of its rectangles where the
 * engine does not show Excel's answer.
 */
import { compare } from '../bench/coverage/compare.js'
import type { RectsAnswer } from '../bench/coverage/case-book.js'
import type { SpillCase } from '../bench/coverage/spill/cases.js'
import { takenBook } from '../bench/coverage/stored.js'
import { compute, InputError, createSession, type BookInput, type ComputeResult } from '../src/index.js'
import { fromLegacy } from '../src/xlsx.js'
import { formatAddress, parseAddress } from '../src/model.js'
import { readDifferences } from './spill-shown.js'

const CLOCK = { now: '2026-10-06T12:00:00Z', seed: 1 }

/** The case's book as a reader loads it, as the engine takes it: the formulas a file stores as of the old kind converted (`fromLegacy`). */
function loaded(c: SpillCase): BookInput {
  const book = takenBook(c.book)
  if (!c.legacy) return book
  const old = new Set(c.legacy)
  return {
    ...book,
    sheets: book.sheets.map((sh) => ({
      ...sh,
      cells: Object.fromEntries(Object.entries(sh.cells).map(([a, x]) => [a, old.has(`${sh.name}!${a}`) && x.formula !== undefined ? { ...x, formula: fromLegacy(x.formula, { names: book.names, sheet: sh.name }) } : x])),
    })),
  }
}

/**
 * The case computed, and its edits applied one by one in a session, as Excel had them typed after computing: an edit
 * of a range edits each of its cells, an edit Excel refuses (part of an old-kind array formula) leaves the book as it was.
 */
export function computed(c: SpillCase): ComputeResult {
  if (!c.edits) return compute(loaded(c), { clock: CLOCK })
  const s = createSession(loaded(c), { clock: CLOCK })
  for (const e of c.edits) {
    const [p, q = p] = e.cell.split(':').map(parseAddress)
    const input = e.value === null ? null : typeof e.value === 'object' ? e.value : { value: e.value }
    const edits = []
    for (let row = p.row; row <= q.row; row++) for (let col = p.col; col <= q.col; col++) edits.push({ sheet: e.sheet, cell: formatAddress(row, col), input })
    try {
      s.setCells(edits)
    } catch (err) {
      if (!(err instanceof InputError)) throw err
    }
  }
  return s.getResult()
}

/** Cells of the case's rectangles where the engine does not show Excel's answer. */
export function differences(c: SpillCase, a: RectsAnswer, r: ComputeResult): string[] {
  return readDifferences(c.read, a.values, r, (excel, engine) => (c.volatile && excel.t === 'n' ? engine.t === 'n' : compare(excel, engine).ok))
}
