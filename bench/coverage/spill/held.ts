/**
 * Held-out spill cases: each answer was predicted from the rules the sets `base` and `order` found
 * (`docs/spill.md`) and committed before Excel computed the case. The rules: an array spills right and down
 * from its anchor, an empty element shows 0; any value (empty text too) or formula in the range, a merged
 * area meeting it or the sheet's edge gives #SPILL!, a format or fill does not; of two spills over one cell
 * the one the other reads wins; `X#` is X's spill range, X alone when X is a formula that did not spill, else
 * #REF!; readers see the spilled values.
 */
import { e, n, one, s, sheet, type SpillCase } from './cases.js'

const seq = (from: number, count: number) => Array.from({ length: count }, (_, i) => n(from + i))
const at = (cells: string[], vals: ReturnType<typeof n>[]) => Object.fromEntries(cells.map((c, i) => [`S!${c}`, vals[i]]))

export function buildSpillCases(): SpillCase[] {
  return [
    one('text-rect', 'an empty text inside a rectangle', { B2: '=SEQUENCE(2,2)', C3: '' }, ['A1:D4'], { predicted: { 'S!B2': e('SPILL'), 'S!C3': s('') } }),
    one('merged-row', 'a row meets a merged area at its end', { A1: '=SEQUENCE(1,4)' }, ['A1:E3'], { merges: ['D1:D2'], predicted: { 'S!A1': e('SPILL') } }),
    one('merged-below', 'a merged area just below a row', { A1: '=SEQUENCE(1,4)' }, ['A1:E4'], { merges: ['A2:D3'], predicted: at(['A1', 'B1', 'C1', 'D1'], seq(1, 4)) }),
    one('ref-rect', 'E2# of a rectangle in functions', { E2: '=SEQUENCE(2,3)', I1: '=SUM(E2#)', I2: '=ROWS(E2#)', I3: '=COLUMNS(E2#)', I4: '=INDEX(E2#,2,3)' }, ['E1:I4'], {
      predicted: { ...at(['E2', 'F2', 'G2', 'E3', 'F3', 'G3'], seq(1, 6)), 'S!I1': n(21), 'S!I2': n(2), 'S!I3': n(3), 'S!I4': n(6) },
    }),
    one('name-row', 'a name over a row spill', { B1: '=SEQUENCE(1,4)', A3: '=COLUMNS(nm)', A4: '=SUM(nm)' }, ['A1:E4'], {
      book: { sheets: [], names: [{ name: 'nm', ref: "'S'!$B$1#" }] },
      predicted: { ...at(['B1', 'C1', 'D1', 'E1'], seq(1, 4)), 'S!A3': n(4), 'S!A4': n(10) },
    }),
    one('col-text', 'a whole column meets an empty text', { A1: 1, A2: 2, B1: '=A:A*2', B5: '' }, ['B1:B5'], { predicted: { 'S!B1': e('SPILL'), 'S!B5': s('') } }),
    one('row-blocked', 'a whole row meets a value', { A1: 1, B1: 2, A2: '=1:1*2', Z2: 5 }, ['A2:B2', 'Z2'], { predicted: { 'S!A2': e('SPILL'), 'S!Z2': n(5) } }),
    one('one-const-ref', 'A1# of a 1×1 constant in arithmetic', { A1: '={7}', B1: '=SUM(A1#)', C1: '=A1#*2' }, ['A1:D2'], { predicted: { 'S!A1': n(7), 'S!B1': n(7), 'S!C1': n(14) } }),
    one('transpose', 'a column transposed spills right', { A1: 1, A2: 2, A3: 3, C1: '=TRANSPOSE(A1:A3)' }, ['C1:F2'], { predicted: at(['C1', 'D1', 'E1'], seq(1, 3)) }),
    one('blocked-chain', 'a spill reference to a blocked anchor', { C1: '=SEQUENCE(3)', C3: 'x', D1: '=C1#', E1: '=ROWS(C1#)' }, ['C1:F4'], {
      predicted: { 'S!C1': e('SPILL'), 'S!C3': s('x'), 'S!D1': e('SPILL'), 'S!E1': n(1) },
    }),
    one('sum-partial', 'ranges over part of a spill and a whole column', { A1: '=SEQUENCE(4)', C1: '=SUM(A2:A3)', C2: '=SUM(A3:A9)', C3: '=COUNT(A:A)' }, ['A1:C5'], {
      predicted: { ...at(['A1', 'A2', 'A3', 'A4'], seq(1, 4)), 'S!C1': n(5), 'S!C2': n(7), 'S!C3': n(4) },
    }),
    {
      id: 'other-sheet-reader',
      about: 'a spilled cell and a whole column read from another sheet',
      book: { sheets: [sheet('S', { A1: '=SEQUENCE(3)*10' }), sheet('T', { B2: '=S!A3+1', B3: '=SUM(S!A:A)' })] },
      read: [{ sheet: 'S', range: 'A1:A4' }, { sheet: 'T', range: 'B1:B4' }],
      predicted: { 'S!A1': n(10), 'S!A2': n(20), 'S!A3': n(30), 'T!B2': n(31), 'T!B3': n(60) },
    },
    one('dep-cross', 'the column reads the row anchor: the row spills', { B1: '=SEQUENCE(4)+0*A3', A3: '=SEQUENCE(1,4)' }, ['A1:E5'], { predicted: { 'S!B1': e('SPILL'), ...at(['A3', 'B3', 'C3', 'D3'], seq(1, 4)) } }),
    one('edge-rect', 'a rectangle ending on the last row and column', { XFC1048575: '=SEQUENCE(2,2)' }, ['XFB1048574:XFD1048576'], { predicted: at(['XFC1048575', 'XFD1048575', 'XFC1048576', 'XFD1048576'], seq(1, 4)) }),
    one('empty-element', 'a lifted IF over a blank cell spills 0 there', { A1: '=IF({1;0;1},"a",B9)' }, ['A1:A4'], { predicted: { 'S!A1': s('a'), 'S!A2': n(0), 'S!A3': s('a') } }),
    one('error-block', 'a formula giving an error in the range', { A1: '=SEQUENCE(3)', A2: '=1/0' }, ['A1:A4'], { predicted: { 'S!A1': e('SPILL'), 'S!A2': e('DIV0') } }),
    one('filter-rect', 'FILTER of two columns spills a rectangle', { A1: 1, A2: 2, A3: 3, A4: 4, B1: 'a', B2: 'b', B3: 'c', B4: 'd', D1: '=FILTER(A1:B4,A1:A4>2)' }, ['D1:E3'], {
      predicted: { 'S!D1': n(3), 'S!E1': s('c'), 'S!D2': n(4), 'S!E2': s('d') },
    }),
  ]
}
