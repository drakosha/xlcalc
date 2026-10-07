/** Renaming a sheet to `History`, the name Microsoft documents as reserved, in other cases and inside longer names. */
import { baseBook, type StructCase, type StructCases } from './cases.js'

export function buildStructCases(): StructCases {
  const cases: StructCase[] = ['History', 'history', 'HISTORY', 'History1', 'My History', 'Histor'].map((name) => ({ id: `rename-S-${name.replace(' ', '-')}`, book: 'base', op: { op: 'renameSheet', sheet: 'S', name } }))
  return { books: { base: baseBook() }, cases }
}
