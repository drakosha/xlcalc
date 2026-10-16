/**
 * Excel's Replace All (`Session.replaceInCells`), measured in Excel for Windows (`bench/coverage/replace/`): the cells
 * Find finds in the text they open with (Look in Formulas, hidden cells skipped as Find skips them), each one's text with
 * the query's matches put as the replacement is (`replacer`), entered back as typed but without the percent editor (`6%`
 * in a percent cell is 0.06, `-5` -500%) or the text format (a number stays a number), as `readEntry` reads text Excel's
 * Replace puts back. Excel enters every cell it found, its text changed or not. A text that
 * reads as something else stays a text, as after an apostrophe (the book keeps no mark of one); an array of the old
 * kind is entered whole from its first cell wherever a cell of it was found. A text Excel would not take (a formula that
 * no longer reads, a formula into a table's header) leaves its cell as it was.
 */
import type { CellRef, CellTextRefusal, RangeRef, Refusal, ReplaceOptions } from '../contract.js'
import type { EngineState } from '../engine.js'
import { allotWrites, type Steps } from '../exec/budget.js'
import { checkString } from '../input.js'
import { inputIn } from '../looks/sheet.js'
import { formatAddress, MAX_ROWS } from '../model.js'
import { ownerOf } from '../spill/place.js'
import { replacer } from '../find-text.js'
import { finderOf } from './find.js'
import { refsOf } from './queries.js'
import { cellText, textRefusal } from './text.js'
import type { Typed } from './texts.js'
import { checkLineNumbers, sessionEditTexts } from './viewport.js'

/** The cells' new inputs, as typed cells, and the cells whose new text Excel would not take. */
export interface Replaced {
  cells: Typed[]
  refused: (Refusal<CellTextRefusal> & CellRef)[]
}

/** The cells Replace All writes over `given` (options checked) at the moment `now`. The steps yield between cells. */
export function* replacedTexts(s: EngineState, given: RangeRef[], query: string, replacement: string, o: ReplaceOptions, now: string): Steps<Replaced> {
  const refs = refsOf(s, given)
  checkString(query, 'query')
  checkString(replacement, 'replacement')
  if (o.rows !== undefined) checkLineNumbers(o.rows, 'opts.rows', MAX_ROWS)
  const finder = finderOf(s, query, { ...o, in: 'edit' }, now)
  const replace = replacer(query, replacement, o)
  const out: Replaced = { cells: [], refused: [] }
  if (!replace) return out
  const editText = sessionEditTexts(s, now)
  // an array of the old kind is entered once, from its first cell
  const arrays = new Set<string>()
  for (const found of finder.found(refs, o.rows?.map((n) => n - 1) ?? null, null)) {
    if (!found) {
      yield
      continue
    }
    const [i, , , , text] = found
    let [, r, c] = found
    let own = inputIn(s.input.sheets[i], formatAddress(r, c), r, c)
    const x = own?.array === undefined && own?.formula === undefined ? ownerOf(s, i, r, c) : undefined
    const top = x && s.input.sheets[i].cells[formatAddress(x.r1, x.c1)]
    if (x && top?.array !== undefined) [r, c, own] = [x.r1, x.c1, top]
    const ref = { sheet: s.sheetNames[i], cell: formatAddress(r, c) }
    if (own?.array !== undefined && arrays.has(`${i}!${ref.cell}`)) continue
    if (own?.array !== undefined) arrays.add(`${i}!${ref.cell}`)
    // a text that reads as something else is entered after the apostrophe it opens with
    const typed = own?.formula === undefined && typeof own?.value === 'string' && editText(i, r, c, own) === `'${text}` ? "'" : ''
    const t = cellText(s, i, r, c, own, typed + (replace(text) ?? text), now, 'replaced')
    // a formula into an array's first cell is the array's; anything else there changes part of it
    const input = t.ok ? (own?.array !== undefined && t.input?.formula !== undefined ? { ...t.input, array: own.array } : t.input) : null
    const refusal = t.ok ? textRefusal(s, { ...ref, input }) : t
    if (refusal) out.refused.push({ ...refusal, ...ref })
    else out.cells.push({ i, r, c, edit: { ...ref, input } })
  }
  // what Find found is known only now: over `limits.writes` the call stops before anything is written
  allotWrites(out.cells.length, s.input.locale)
  return out
}
