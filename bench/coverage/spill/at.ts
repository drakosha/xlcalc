/**
 * The implicit intersection operator `@` in dynamic array formulas (`run.ts spill at`; a file writes it
 * `_xlfn.SINGLE(...)`): a range's cell in the formula's row or column, on another sheet, outside the range, whole
 * columns and rows, a spill reference, an array, a call answering a reference or an array, a name, a union; and
 * whether its answer is a reference (SUM skips text there) or a value.
 */
import { one, sheet, type CellSpec, type SpillCase } from './cases.js'

/** Data on S: A1:A10 1…10 with text "5" in A5, C1:E1 a row. */
const DATA: Record<string, CellSpec> = { ...Object.fromEntries(Array.from({ length: 10 }, (_, i) => [`A${i + 1}`, i + 1])), A5: '5', C1: 7, D1: 8, E1: 9 }

export function buildSpillCases(): SpillCase[] {
  return [
    one('column', "a column: the formula's row, else #VALUE!", { ...DATA, G3: '=@A1:A10', G12: '=@A1:A10', G4: '=@A1:A10*2', G6: '=@A:A', G7: '=@$A$1:$A$10+1', G8: '=@A1:A10+@A1:A10' }, ['G1:G12']),
    one('row', "a row: the formula's column", { ...DATA, D3: '=@C1:E1', H3: '=@C1:E1', C4: '=@1:1', H5: '=@C:E' }, ['C3:H5']),
    {
      id: 'other-sheet',
      about: 'a range of another sheet by the same row and column, a block',
      book: { sheets: [sheet('S', { ...DATA, G1: 1, H1: 2, G2: 3, H2: 4 }), sheet('T', { B3: '=@S!A1:A10', B11: '=@S!A1:A10', G2: '=@S!G1:H2', H1: '=@S!G1:H2', J2: '=@S!G1:H2', G5: '=@S!G1:H2' })] },
      read: [{ sheet: 'T', range: 'B1:J11' }],
    },
    one('arrays', 'an array: its top-left value; a function with an array answer', { ...DATA, H1: '=@{1,2,3}', H2: '=@{4;5}*2', H3: '=@SEQUENCE(3)', H4: '=@(A1:A3*10)', H5: '=@SEQUENCE(2,2)+1', H6: '=@5', H7: '=@"a"', H8: '=@TRUE' }, ['H1:I9']),
    one('calls', 'a call answering a reference, a spill reference, a single cell', { ...DATA, C3: '=SEQUENCE(4)', H2: '=@INDEX(A1:B10,0,1)', H4: '=@C3#', H9: '=@C3#', H5: '=@A2', H6: '=@IF(TRUE,A1:A10)', H7: '=@CHOOSE(1,A1:A10)', H8: '=@XLOOKUP(3,A1:A10,A1:A10)' }, ['C3:C6', 'H2:H9']),
    one('reference', 'the answer is a reference: SUM, ROW and ISREF of it, text in the cell', { ...DATA, H5: '=SUM(@A1:A10)', H6: '=SUM(@A1:A10)', I5: '=ROW(@A1:A10)', I6: '=ISREF(@A1:A10)', I7: '=COUNT(@A1:A10)', I8: '=COUNTA(@A1:A10)', I9: '=@A1:A10&""' }, ['H5:I9']),
    one('names', 'a name over a range, an array, a union', { ...DATA, H3: '=@nr', H4: '=@na', H5: '=@nu', H6: '=SUM(@nr)' }, ['H3:H6'], {
      book: {
        sheets: [],
        names: [
          { name: 'nr', ref: "'S'!$A$1:$A$10" },
          { name: 'na', ref: '{4,5,6}' },
          { name: 'nu', ref: "'S'!$A$1:$A$3,'S'!$C$1:$C$3" },
        ],
      },
    }),
    one('union', 'a union and a reference across sheets under @', { ...DATA, H2: '=@(A1:A3,C1:C3)' }, ['H2'], {
      book: { sheets: [sheet('T', { A2: 1 }), sheet('U', { A2: 2 }), sheet('V', { B2: '=@T:U!A1:A3', B3: '=@T:U!A2' })] },
      read: [{ sheet: 'S', range: 'H2' }, { sheet: 'V', range: 'B2:B3' }],
    }),
  ]
}
