/** Inserting and deleting rows and columns of the base book: inside, at the edges and around every kind of reference. */
import { baseBook, type StructCase, type StructCases } from './cases.js'

export function buildStructCases(): StructCases {
  const cases: StructCase[] = []
  const rows = (id: string, op: 'insertRows' | 'deleteRows', sheet: string, row: number, count = 1) => cases.push({ id, book: 'base', op: { op, sheet, row, count } })
  const cols = (id: string, op: 'insertColumns' | 'deleteColumns', sheet: string, column: string, count = 1) => cases.push({ id, book: 'base', op: { op, sheet, column, count } })

  rows('insert-rows-5x2', 'insertRows', 'S', 5, 2)
  rows('insert-rows-3', 'insertRows', 'S', 3)
  rows('insert-rows-9', 'insertRows', 'S', 9)
  rows('insert-rows-1', 'insertRows', 'S', 1)
  rows('insert-rows-30', 'insertRows', 'S', 30)
  rows('insert-rows-31', 'insertRows', 'S', 31)
  rows('delete-rows-5x2', 'deleteRows', 'S', 5, 2)
  rows('delete-rows-6', 'deleteRows', 'S', 6)
  rows('delete-rows-3x6', 'deleteRows', 'S', 3, 6)
  rows('delete-rows-1x4', 'deleteRows', 'S', 1, 4)
  rows('delete-rows-30', 'deleteRows', 'S', 30)
  rows('delete-rows-30x2', 'deleteRows', 'S', 30, 2)
  rows('delete-rows-12', 'deleteRows', 'S', 12)
  rows('insert-rows-T1', 'insertRows', 'T', 1)
  rows('delete-rows-T5', 'deleteRows', 'T', 5)

  cols('insert-cols-B', 'insertColumns', 'S', 'B')
  cols('insert-cols-Ax2', 'insertColumns', 'S', 'A', 2)
  cols('insert-cols-C', 'insertColumns', 'S', 'C')
  cols('insert-cols-E', 'insertColumns', 'S', 'E')
  cols('delete-cols-B', 'deleteColumns', 'S', 'B')
  cols('delete-cols-A', 'deleteColumns', 'S', 'A')
  cols('delete-cols-Bx2', 'deleteColumns', 'S', 'B', 2)
  cols('delete-cols-D', 'deleteColumns', 'S', 'D')
  cols('delete-cols-E', 'deleteColumns', 'S', 'E')
  cols('insert-cols-T-A', 'insertColumns', 'T', 'A')
  cols('delete-cols-MyData-A', 'deleteColumns', 'My Data', 'A')
  return { books: { base: baseBook() }, cases }
}
