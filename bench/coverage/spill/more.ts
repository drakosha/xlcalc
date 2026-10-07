/**
 * Spilling not measured by the first sets: an array whose size comes from a volatile draw (`SEQUENCE(RANDBETWEEN(1,3))`:
 * does Excel give #SPILL! for a size that changes between its passes), and cells typed into and cleared from a
 * spill range after the computation, in books with merged areas next to the range, inside it or holding the anchor.
 */
import { one, type SpillCase } from './cases.js'

/** Ten independent arrays of a drawn size, two columns apart. */
const drawn = Object.fromEntries(Array.from({ length: 10 }, (_, i) => [`${String.fromCharCode(65 + 2 * i)}1`, '=SEQUENCE(RANDBETWEEN(1,3))']))

const typed = (cell: string, value: number | string | null) => ({ sheet: 'S', cell, value })

export function buildSpillCases(): SpillCase[] {
  return [
    one('volatile-size', 'the size of a column from a draw', { A1: '=SEQUENCE(RANDBETWEEN(1,3))' }, ['A1:A4'], { volatile: true }),
    one('volatile-sizes', 'ten columns whose sizes are drawn', drawn, ['A1:T4'], { volatile: true }),
    one('volatile-fixed', 'a drawn size that cannot change', { A1: '=SEQUENCE(RANDBETWEEN(2,2))' }, ['A1:A3'], { volatile: true }),
    one('volatile-cell', 'the size from a cell holding a draw', { B1: '=RANDBETWEEN(1,3)', A1: '=SEQUENCE(B1)' }, ['A1:B4'], { volatile: true }),
    one('volatile-randarray', 'RANDARRAY of a drawn size', { A1: '=RANDARRAY(RANDBETWEEN(1,3))' }, ['A1:A4'], { volatile: true }),
    one('edit-typed', 'a value typed into the range, a merged area beside it', { A1: '=SEQUENCE(3)', B5: 1 }, ['A1:D5'], { merges: ['C1:D2'], edits: [typed('A2', 5)] }),
    one('edit-cleared', 'a value typed into the range and cleared, a merged area beside it', { A1: '=SEQUENCE(3)' }, ['A1:D5'], { merges: ['C1:D2'], edits: [typed('A2', 5), typed('A2', null)] }),
    one('edit-text', 'text typed into a rectangle, a merged area below it', { A1: '=SEQUENCE(3,2)' }, ['A1:C6'], { merges: ['A5:B6'], edits: [typed('B2', 'x')] }),
    one('edit-text-cleared', 'text typed into a rectangle and cleared, a merged area below it', { A1: '=SEQUENCE(3,2)' }, ['A1:C6'], { merges: ['A5:B6'], edits: [typed('B2', 'x'), typed('B2', null)] }),
    one('edit-merged-block', 'the range meets a merged area; a value typed in and cleared', { A1: '=SEQUENCE(3)' }, ['A1:B5'], { merges: ['A3:B4'], edits: [typed('A2', 1), typed('A2', null)] }),
    one('edit-value-cleared', 'a value in the range cleared, a merged area beside it', { A1: '=SEQUENCE(3)', A3: 7 }, ['A1:C4'], { merges: ['B1:C1'], edits: [typed('A3', null)] }),
    one('edit-reader', 'a value typed into a range that a reader sums', { A1: '=SEQUENCE(3)', C1: '=SUM(A1#)', D1: '=A1#*2' }, ['A1:E4'], { merges: ['E1:E2'], edits: [typed('A3', 9)] }),
    one('edit-anchor-merged', 'the anchor holds a merged area; a value typed beside it', { A1: '=SEQUENCE(1,3)' }, ['A1:D3'], { merges: ['A1:A2'], edits: [typed('D1', 1)] }),
  ]
}
