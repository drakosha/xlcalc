/**
 * The first spill set: directions, what blocks a spill, two spills over one cell, the sheet's edges, spill
 * references (`E2#`) and their readers, sizes that come from data, whole columns and rows, arrays of one value
 * and volatile arrays.
 */
import { one, sheet, type SpillCase } from './cases.js'

const DATA6 = { A1: 5, A2: 1, A3: 7, A4: 2, A5: 9, A6: 3 }

export function buildSpillCases(): SpillCase[] {
  return [
    // directions and what spills
    one('down', 'a column spills down', { A1: '=SEQUENCE(3)' }, ['A1:B4']),
    one('right', 'a row spills right', { A1: '=SEQUENCE(1,3)' }, ['A1:D2']),
    one('rect', 'a rectangle spills right and down', { B2: '=SEQUENCE(2,3,10)' }, ['A1:E4']),
    one('range', 'a reference of several cells spills its values; a blank cell', { A1: 1, A3: 'x', A4: true, C1: '=A1:A4', E1: '=ISBLANK(C2)', E2: '=C2&"|"', E3: '=COUNTA(C1:C4)', E4: '=COUNTBLANK(C1:C4)' }, ['A1:E5']),
    one('constant', 'an array constant with text, a boolean and an error', { A1: '={1,"a";TRUE,#N/A}' }, ['A1:C3']),
    one('one-seq', 'an array of one value: the value; its spill reference', { A1: '=SEQUENCE(1,1)', B1: '=A1#', C1: '=ROWS(A1#)', D1: '=SUM(A1#)' }, ['A1:E2']),
    one('one-const', 'a 1×1 constant', { A1: '={5}', B1: '=A1#', C1: '=ROWS(A1#)' }, ['A1:D2']),

    // what blocks a spill
    one('block-value', 'a value in the spill range', { A1: '=SEQUENCE(3)', A3: 5 }, ['A1:A4']),
    one('block-formula', 'a formula in the spill range', { A1: '=SEQUENCE(3)', A3: '=1+1' }, ['A1:A4']),
    one('block-formula-text', 'a formula giving "" in the spill range', { A1: '=SEQUENCE(3)', A3: '=""' }, ['A1:A4']),
    one('block-text', 'a constant empty text in the spill range', { A1: '=SEQUENCE(3)', A3: '' }, ['A1:A4']),
    one('block-format', 'an empty cell with a number format in the spill range', { A1: '=SEQUENCE(3)', A3: { format: '0.00' } }, ['A1:A4']),
    one('block-fill', 'an empty cell with a fill in the spill range', { A1: '=SEQUENCE(3)', A3: { style: 0 } }, ['A1:A4'], { book: { sheets: [], styles: [{ fill: { color: '#FFFF00' } }] } }),
    one('block-corner', 'a value in the far corner of a rectangle', { B2: '=SEQUENCE(3,3)', D4: 7 }, ['A1:E5']),
    one('block-dependent', 'a formula reading the anchor blocks it', { A1: '=SEQUENCE(3)', A3: '=A1+1' }, ['A1:A4']),
    one('merged-inside', 'a merged area inside the spill range', { A2: '=SEQUENCE(1,3)' }, ['A1:D3'], { merges: ['B2:C2'] }),
    one('merged-anchor', 'the anchor is the top-left cell of a merged area', { A1: '=SEQUENCE(3)' }, ['A1:B4'], { merges: ['A1:B1'] }),
    one('merged-partial', 'a merged area that the spill range cuts', { A1: '=SEQUENCE(3)' }, ['A1:B5'], { merges: ['A3:A4'] }),
    one('merged-beside', 'a merged area next to the spill range', { A1: '=SEQUENCE(3)' }, ['A1:C4'], { merges: ['B2:C3'] }),

    // two spills over one cell
    one('cross', 'a column spill and a row spill crossing at C2', { C1: '=SEQUENCE(3)', A2: '=SEQUENCE(1,3)' }, ['A1:D4']),
    one('cross-row-after', 'the row spill reads the column anchor', { C1: '=SEQUENCE(3)', A2: '=SEQUENCE(1,3)+0*C1' }, ['A1:D4']),
    one('cross-col-after', 'the column spill reads the row anchor', { C1: '=SEQUENCE(3)+0*A2', A2: '=SEQUENCE(1,3)' }, ['A1:D4']),
    one('cross-low', 'a longer column and a lower row crossing at B3', { B1: '=SEQUENCE(4)', A3: '=SEQUENCE(1,4)' }, ['A1:E5']),
    one('cross-rect', 'two rectangles overlapping in B2:C3', { B1: '=SEQUENCE(4,2)', A2: '=SEQUENCE(2,4)' }, ['A1:E5']),
    one('chain', 'a spill of a spill', { A1: '=SEQUENCE(3)', B1: '=A1#*10', C1: '=SUM(B1#)' }, ['A1:D4']),

    // the edges of the sheet
    one('edge-down', 'three rows from the second-to-last row', { A1048575: '=SEQUENCE(3)' }, ['A1048573:A1048576']),
    one('edge-down-fit', 'three rows ending on the last row', { A1048574: '=SEQUENCE(3)' }, ['A1048573:A1048576']),
    one('edge-right', 'three columns from the second-to-last column', { XFC1: '=SEQUENCE(1,3)' }, ['XFA1:XFD2']),
    one('edge-right-fit', 'three columns ending on the last column', { XFB1: '=SEQUENCE(1,3)' }, ['XFA1:XFD2']),

    // spill references
    one('ref', 'E2# in a reference, in functions and spilled itself', { E2: '=SEQUENCE(3)', G1: '=E2#', H1: '=SUM(E2#)', H2: '=ROWS(E2#)', H3: '=INDEX(E2#,2)', H4: '=COLUMNS(E2#)', I1: '=E2#*2' }, ['E1:I5']),
    one('ref-blocked', 'E2# of an anchor that gives #SPILL!', { E2: '=SEQUENCE(3)', E4: 5, G1: '=E2#', G2: '=SUM(E2#)', G3: '=ROWS(E2#)' }, ['E1:G5']),
    one('ref-inside', 'E3# of a cell inside a spill range', { E2: '=SEQUENCE(3)', G1: '=E3#', G2: '=ROWS(E3#)' }, ['E1:G4']),
    one('ref-value', 'A1# of a value', { A1: 5, B1: '=A1#', C1: '=SUM(A1#)' }, ['A1:D2']),
    one('ref-scalar', 'A1# of a formula with one value', { A1: '=1+1', B1: '=A1#', C1: '=ROWS(A1#)' }, ['A1:D2']),
    one('ref-empty', 'A1# of an empty cell', { B1: '=A1#', C1: '=ROWS(A1#)' }, ['A1:D2']),
    one('ref-name', 'a name defined as a spill reference', { E2: '=SEQUENCE(3)', G1: '=SUM(sp)', G2: '=ROWS(sp)', G3: '=sp' }, ['E1:G6'], { book: { sheets: [], names: [{ name: 'sp', ref: "'S'!$E$2#" }] } }),
    one('ref-sheet', 'a spill reference and a spilled cell read from another sheet', { E2: '=SEQUENCE(3)' }, ['E1:E5'], {
      book: { sheets: [sheet('T', { A1: '=SUM(S!E2#)', A2: '=S!E2#', B1: '=S!E3*10' })] },
      read: [{ sheet: 'S', range: 'E1:E5' }, { sheet: 'T', range: 'A1:B5' }],
    }),
    one('readers', 'formulas reading cells inside a spill range', { E2: '=SEQUENCE(3)', G1: '=E3*10', G2: '=SUM(E1:E5)', G3: '=E5', G4: '=COUNT(E:E)', G5: '=ISBLANK(E3)', G6: '=E4&"|"' }, ['E1:G7']),

    // sizes from data
    one('size-fits', 'SEQUENCE sized by a cell: room for 2', { B1: 2, A1: '=SEQUENCE(B1)', A4: 'x' }, ['A1:B5']),
    one('size-blocked', 'SEQUENCE sized by a cell: 4 rows meet a value', { B1: 4, A1: '=SEQUENCE(B1)', A4: 'x' }, ['A1:B5']),
    one('filter', 'FILTER spills the kept values', { ...DATA6, C1: '=FILTER(A1:A6,A1:A6>2)' }, ['C1:C7']),
    one('filter-none', 'FILTER that keeps nothing', { ...DATA6, C1: '=FILTER(A1:A6,A1:A6>100)' }, ['C1:C3']),
    one('unique', 'UNIQUE spills the distinct values', { A1: 'a', A2: 'b', A3: 'a', A4: 'c', A5: 'b', A6: 'a', C1: '=UNIQUE(A1:A6)' }, ['C1:C5']),
    one('sort', 'SORT spills the sorted values', { ...DATA6, C1: '=SORT(A1:A6)' }, ['C1:C7']),

    // whole columns and rows
    one('col-free', 'a whole column spills over a free column', { A1: 1, A2: 2, A3: 3, B1: '=A:A*2', D1: '=SUM(B:B)', D2: '=COUNT(B:B)', D3: '=COUNTA(B:B)', D4: '=B1048576', D5: '=COUNTBLANK(B:B)', D6: '=ROWS(B1#)' }, ['B1:B5', 'B1048574:B1048576', 'D1:D6']),
    one('col-blocked', 'a whole column meets a value far below', { A1: 1, A2: 2, A3: 3, B1: '=A:A*2', B100: 'x' }, ['B1:B3', 'B100']),
    one('col-row2', 'a whole column from row 2 does not fit', { A1: 1, A2: 2, A3: 3, B2: '=A:A*2' }, ['B1:B4']),
    one('col-format', 'a whole column over a formatted empty cell', { A1: 1, A2: 2, A3: 3, B1: '=A:A*2', B700: { format: '0.00' } }, ['B1:B3', 'B700']),
    one('row-free', 'a whole row spills right', { A1: 1, B1: 2, C1: 3, A2: '=1:1*2' }, ['A2:E2', 'XFC2:XFD2']),

    // volatile
    one('rand', 'RANDARRAY spills; its size is fixed', { A1: '=RANDARRAY(3)', B1: '=ROWS(A1#)', B2: '=AND(A1#>=0,A1#<1)' }, ['A1:B4'], { volatile: true }),
  ]
}
