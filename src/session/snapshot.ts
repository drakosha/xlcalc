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
    let addresses: string[] = []
    let chunk: CellInput[] = []
    // Clone cell arrays, keeping address keys out of the structured clone's intermediate objects.
    const flush = () => {
      const values = clone(chunk)
      const count = addresses.length
      for (let i = 0; i < count; i++) copied.cells[addresses[i]] = values[i]
      addresses = []
      chunk = []
      return paced(count)
    }
    for (const addr in cells) {
      addresses.push(addr)
      chunk.push(cells[addr])
      if (chunk.length === CHUNK && flush()) yield
    }
    if (chunk.length && flush()) yield
    out.sheets.push(copied)
  }
  return out
}
