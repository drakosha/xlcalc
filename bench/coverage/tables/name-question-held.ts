/**
 * Held out: a file whose names hold `?` just before what the reader of a file's formulas looks at, as `writeXlsx` writes
 * it: defined names with `?` before `_xlfn.`, `_xlpm.` and `_xludf.` (`a?_xlfn.b`), which the reader took for a prefix
 * to drop, and a table named `T?`, whose structured reference `T?[x]` it took for a link to another book. Excel loads
 * the book and computes it; `test/name-question.test.ts` reads the same file back. Predicted by the engine before Excel
 * ran it (`name-question-held.predicted.json`).
 */
import type { NameInput } from '../../../src/contract.js'
import { book, on, table, type TableCase } from './cases.js'

const names: NameInput[] = [
  { name: 'a?_xlfn.b', ref: '5' },
  { name: 'x?_xlpm.y', ref: '7' },
  { name: 'u?_xludf.v', ref: '11' },
]

export function buildTableCases(): TableCase[] {
  const t = table('T?', 'A1', ['x'], [[2], [3]])
  return [
    {
      id: 'names',
      about: 'names with ? before a prefix of a newer function, a LET name and a user function, in formulas',
      book: { ...book(['S', { C1: '=a?_xlfn.b+1', C2: '=x?_xlpm.y*2', C3: '=u?_xludf.v+a?_xlfn.b' }]), names },
      read: on('S', 'C1:C3'),
    },
    {
      id: 'table',
      about: 'a table named T? and structured references to it',
      book: book(['S', { C1: '=SUM(T?[x])', C2: '=ROWS(T?[#Data])', C3: '=SUM(T?[[#Data],[x]])*10' }, t]),
      read: on('S', 'A1:C3'),
    },
  ]
}
