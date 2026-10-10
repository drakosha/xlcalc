/**
 * Threads of comments (Excel 365's comments), put into the file by hand as Excel 365 writes them (`threaded.ts`): a
 * thread of three comments by two people, one of one comment, one on another sheet, a note beside them; the same with a
 * row inserted above. The reader takes a thread as a note: its first author, every comment's text, the replies after an
 * empty line with their authors. Predicted before Excel ran it (`threads.predicted.json`).
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import type { PartsCases, ThreadSpec } from './cases.js'

function plainBook(): BookInput {
  const cells: Record<string, CellInput> = { B4: { value: 'b4', note: { author: 'Ann', text: 'A note beside the threads' } } }
  for (let r = 1; r <= 6; r++) Object.assign(cells, { [`A${r}`]: { value: r }, [`B${r}`]: { value: `b${r}`, ...cells[`B${r}`] } })
  return { sheets: [{ name: 'S', cells }, { name: 'T', cells: { A1: { value: 1 } } }] }
}

const threads: ThreadSpec[] = [
  { sheet: 'S', cell: 'B2', comments: [{ author: 'Ann', text: 'First comment' }, { author: 'Борис', text: 'A reply' }, { author: 'Ann', text: 'Another reply\non two lines' }] },
  { sheet: 'S', cell: 'C3', comments: [{ author: 'Борис', text: 'Single' }] },
  { sheet: 'T', cell: 'A1', comments: [{ author: 'Ann', text: 'On T' }, { author: 'Борис', text: 'Ответ' }] },
]

export function buildPartsCases(): PartsCases {
  return {
    books: { plain: plainBook() },
    cases: [
      { id: 'threads', book: 'plain' },
      { id: 'threads-insert-row', book: 'plain', op: { op: 'insertRows', sheet: 'S', row: 1 } },
    ],
    threads: { plain: threads },
  }
}
