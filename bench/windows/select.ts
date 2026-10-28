/**
 * Selections in Excel for Windows (`bench/coverage/select/`): each book of a set written by `writeXlsx` and opened, then
 * each case's `Range.CreateNames` over its range with its sides of labels, the names Excel made (deleted before the next
 * case), or its `Range.SpecialCells`, the rectangles Excel selected. The cases of one book run in one file. The answers go to
 * `test/excel/win/select/<set>.json`; a case Excel refuses keeps its error.
 */
import type { SpecialCells, SpecialValue } from '../../src/contract.js'
import { cellsAnswer, namesAnswer, selectBook, selectFile, type SelectBook, type SelectSet } from '../coverage/select/cases.js'
import { setModule } from '../coverage/stored.js'
import type { Step, WinEnv } from './remote.js'
import { stepByBook } from './by-book.js'

/** Excel's `XlCellType` and `XlSpecialCellsValue`. */
const XL_TYPE: Record<SpecialCells['type'], number> = { blanks: 4, constants: 2, formulas: -4123, visible: 12 }
const XL_VALUE: Record<SpecialValue, number> = { numbers: 1, text: 2, logicals: 4, errors: 16 }

/** Go To Special's step: the type, and the kinds of values as their sum (all of them without any). */
function specialStep(range: string, k: SpecialCells): Step {
  const values = 'values' in k && k.values ? k.values.reduce((n, v) => n + XL_VALUE[v], 0) : 0
  return { do: 'special', sheet: 'S', range, type: XL_TYPE[k.type], values }
}

export async function stepSelect(set: string, e: WinEnv): Promise<void> {
  const { BOOKS, buildSelectCases } = await setModule<SelectSet>('select', set)
  // a book without a filter or typed texts
  const books: Record<string, SelectBook & { typed?: never }> = BOOKS
  await stepByBook({
    kind: 'select',
    set,
    tag: 'en-US',
    e,
    books,
    cases: buildSelectCases(),
    file: selectFile(set),
    sent: selectBook,
    step: (c) => ('labels' in c ? { do: 'createNames', sheet: 'S', range: c.range, ...c.labels } : specialStep(c.range, c.special)),
    answer: (x) => {
      if (x && typeof x === 'object' && 'error' in x) return { error: String(x.error) }
      if (!(x && typeof x === 'object' && 'areas' in x)) return namesAnswer(x as [string, string][])
      const { areas, count } = x as { areas: string[]; count?: number }
      // an address cut short (COM gives at most 255 characters of it) is an error, not cells
      return count === undefined || count === areas.length ? cellsAnswer(areas) : { error: `${count} areas, ${areas.length} in the address` }
    },
  })
}
