/**
 * Held-out checks of the style and layout rules fitted on the set `styles` (predicted by the engine before
 * Excel ran): spans of rows and columns under inserts and deletes, tiles and parts of merged areas in copies,
 * merged areas moved to and landed on another sheet, whole rows and columns moved and copied in tiles, the
 * looks of cells moved into styled rows and out of styled columns; last, a copy of a column landing on part
 * of a merged area (refused).
 */
import type { BookInput, RangeRef } from '../../../src/contract.js'
import type { StructCase, StructCases } from './cases.js'
import { styledBook } from './styles.js'

/** The styled book with columns F:G 10 wide, rows 12:13 at height 20, B2:C2 merged; T with B2:C3 merged. */
function checksBook(): BookInput {
  const b = styledBook()
  const [s, t] = b.sheets
  return {
    ...b,
    sheets: [
      { ...s, columns: { ...s.columns, 'F:G': { width: 10 } }, rows: { ...s.rows, '12:13': { height: 20 } }, merges: [...s.merges!, 'B2:C2'] },
      { ...t, merges: ['B2:C3'] },
    ],
  }
}

export function buildStructCases(): StructCases {
  const cases: StructCase[] = []
  const at = (ref: string): RangeRef => {
    const i = ref.lastIndexOf('!')
    return { sheet: ref.slice(0, i), range: ref.slice(i + 1) }
  }
  const op = (id: string, o: StructCase['op']) => cases.push({ id, book: 'checks', op: o })
  const move = (from: string, to: string) => {
    const t = at(to)
    op(`move-${from}-${to}`, { op: 'moveRange', from: at(from), to: { sheet: t.sheet, cell: t.range } })
  }
  const copy = (from: string, to: string) => op(`copy-${from}-${to}`, { op: 'copyRange', from: at(from), to: at(to) })

  op('insert-cols-Fx2', { op: 'insertColumns', sheet: 'S', column: 'F', count: 2 }) // right of E, bold; F:G moves
  op('insert-rows-12x2', { op: 'insertRows', sheet: 'S', row: 12, count: 2 }) // at the first row of 12:13
  op('insert-rows-13', { op: 'insertRows', sheet: 'S', row: 13 }) // inside 12:13
  op('insert-rows-16x2', { op: 'insertRows', sheet: 'S', row: 16, count: 2 }) // inside A14:A16
  op('delete-rows-13', { op: 'deleteRows', sheet: 'S', row: 13 })
  op('delete-cols-F', { op: 'deleteColumns', sheet: 'S', column: 'F' })
  copy('S!B2:C2', 'S!E5:F7') // a merged area tiled
  copy('S!A19:B20', 'S!B10:C11') // onto the whole of B10:C11
  copy('S!A14', 'S!E14') // one cell of A14:A16
  move('S!B10:C11', 'T!E5') // a merged area to another sheet
  move('S!A19:A20', 'T!C3') // onto part of T's B2:C3
  move('S!8:9', 'S!A20') // rows at height 24 and orange
  move('S!H:H', 'S!J1') // the orange column onto part of I2:J2
  move('S!E1:E3', 'S!B20') // out of the bold column
  move('S!B3:D3', 'S!B6') // into the blue row
  copy('S!A6:C6', 'S!E20') // the blue row's cells into the bold column
  copy('S!6:6', 'S!20:22') // the blue row in tiles
  // last: a copy landing on part of a merged area (an error, if no window)
  copy('S!C:C', 'S!I1')
  return { books: { checks: checksBook() }, cases }
}
