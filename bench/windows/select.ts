/**
 * Create Names from Selection in Excel for Windows (`bench/coverage/select/`): each book of a set written by
 * `writeXlsx` and opened, then each case's `Range.CreateNames` over its range with its sides of labels, the names Excel
 * made (deleted before the next case). The cases of one book run in one file. The answers go to
 * `test/excel/win/select/<set>.json`; a case Excel refuses keeps its error.
 */
import { namesAnswer, selectBook, selectFile, type SelectBook, type SelectSet } from '../coverage/select/cases.js'
import { setModule } from '../coverage/stored.js'
import type { WinEnv } from './remote.js'
import { stepByBook } from './by-book.js'

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
    step: (c) => ({ do: 'createNames', sheet: 'S', range: c.range, ...c.labels }),
    answer: (x) => (x && typeof x === 'object' && 'error' in x ? { error: String(x.error) } : namesAnswer(x as [string, string][])),
  })
}
