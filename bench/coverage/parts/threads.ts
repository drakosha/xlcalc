/**
 * Threads of comments (Excel 365's comments), which Excel for Windows makes through its object model
 * (`AddCommentThreaded`, `AddReply`) and Excel for Mac does not: a thread of three comments, one of one, a note beside
 * them, and the same with a row inserted above. The reader takes a thread as a note: its first author, every comment's
 * text, the replies after an empty line with their authors. Predicted before Excel ran it (`threads.predicted.json`).
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import type { PartsCases, PartsMake } from './cases.js'

function plainBook(): BookInput {
  const cells: Record<string, CellInput> = {}
  for (let r = 1; r <= 6; r++) Object.assign(cells, { [`A${r}`]: { value: r }, [`B${r}`]: { value: `b${r}` } })
  return { sheets: [{ name: 'S', cells }, { name: 'T', cells: { A1: { value: 1 } } }] }
}

const make: PartsMake[] = [
  { thread: 'B2', sheet: 'S', texts: ['First comment', 'A reply', 'Another reply\non two lines'] },
  { thread: 'C3', sheet: 'S', texts: ['Single'] },
  { note: 'B4', sheet: 'S', text: 'A note beside the threads' },
  { thread: 'A1', sheet: 'T', texts: ['On T', 'Ответ'] },
]

export function buildPartsCases(): PartsCases {
  return {
    books: { plain: plainBook() },
    cases: [
      { id: 'made-threads', book: 'plain', make },
      { id: 'threads-insert-row', book: 'plain', make, op: { op: 'insertRows', sheet: 'S', row: 1 } },
    ],
    windows: true,
  }
}
