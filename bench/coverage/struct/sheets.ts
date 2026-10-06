/** Renaming and deleting sheets of the base book, and references that are already `#REF!`. */
import type { BookInput } from '../../../src/contract.js'
import { baseBook, type StructCase, type StructCases } from './cases.js'

/** Formulas Excel left after deleting rows and a sheet: a sheet with an error reference, an error with a reference. */
function brokenBook(): BookInput {
  return {
    sheets: [
      { name: 'S', cells: { A1: { value: 1 } } },
      { name: 'T', cells: { A1: { formula: 'S!#REF!' }, A2: { formula: 'SUM(S!#REF!,1)' }, A3: { formula: 'S!A1+1' }, A4: { formula: '#REF!+1' } } },
      { name: 'U', cells: { A1: { value: 2 } } },
    ],
    names: [{ name: 'gone', ref: "'S'!#REF!" }],
  }
}

export function buildStructCases(): StructCases {
  const cases: StructCase[] = []
  const rename = (id: string, sheet: string, name: string, book = 'base') => cases.push({ id, book, op: { op: 'renameSheet', sheet, name } })
  const drop = (id: string, sheet: string, book = 'base') => cases.push({ id, book, op: { op: 'deleteSheet', sheet } })

  rename('rename-S-quoted', 'S', 'New Name')
  rename('rename-S-case', 'S', 's')
  rename('rename-MyData', 'My Data', 'Data2')
  rename('rename-T-apostrophe', 'T', "It's")
  drop('delete-S', 'S')
  drop('delete-T', 'T')
  drop('delete-MyData', 'My Data')
  rename('broken-rename-S', 'S', 'V', 'broken')
  drop('broken-delete-S', 'S', 'broken')
  return { books: { base: baseBook(), broken: brokenBook() }, cases }
}
