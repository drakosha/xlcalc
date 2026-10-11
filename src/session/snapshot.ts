/** Copies a snapshot's book in portions, so inline calls can yield and stop while copying a large sheet. */
import type { BookInput, CellInput } from '../contract.js'
import { paced, type Steps } from '../exec/budget.js'

const clone = (globalThis as unknown as { structuredClone<T>(value: T): T }).structuredClone
const CHUNK = 1024

/** A detached book, including nested styles, layout and cell metadata; numeric values are preserved as given. */
export function* snapshotBook(book: BookInput): Steps<BookInput> {
  const { sheets, ...head } = book
  const out: BookInput = { ...clone(head), sheets: [] }
  for (const sheet of sheets) {
    const { cells, ...rest } = sheet
    const copied = { ...clone(rest), cells: {} as Record<string, CellInput> }
    let chunk: Record<string, CellInput> = {}
    let count = 0
    for (const addr in cells) {
      chunk[addr] = cells[addr]
      if (++count === CHUNK) {
        Object.assign(copied.cells, clone(chunk))
        chunk = {}
        count = 0
        if (paced(CHUNK)) yield
      }
    }
    if (count) {
      Object.assign(copied.cells, clone(chunk))
      if (paced(count)) yield
    }
    out.sheets.push(copied)
  }
  return out
}
