/**
 * More spill references under structure operations, on the book of `spill.ts`: copies (a relative and an
 * absolute `E2#`), renaming and deleting the anchors' sheet (`#REF!E2#`). A set of its own because Excel
 * stopped on a recovery window (06.10) before they ran with `spill.ts`; they were computed on 07.10.
 */
import { spillBook } from './spill.js'
import type { StructCase, StructCases } from './cases.js'

export function buildStructCases(): StructCases {
  const cases: StructCase[] = []
  const add = (id: string, op: StructCase['op']) => cases.push({ id, book: 'spill', op })
  add('copy-G1-I1', { op: 'copyRange', from: { sheet: 'S', range: 'G1' }, to: { sheet: 'S', range: 'I1' } })
  add('copy-G6-I6', { op: 'copyRange', from: { sheet: 'S', range: 'G6' }, to: { sheet: 'S', range: 'I6' } })
  add('rename-S', { op: 'renameSheet', sheet: 'S', name: 'R' })
  add('delete-S', { op: 'deleteSheet', sheet: 'S' })
  return { books: { spill: spillBook() }, cases, dynamic: true }
}
