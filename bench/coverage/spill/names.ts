/**
 * Defined names of any formula (`run.ts spill names`): constants, formulas over references, a reference answered by
 * INDEX, unions, references without `$` (relative to A1, moved by the cell that uses the name), ROW and COLUMN in a
 * name, names local to a sheet and names qualified by a sheet, a name's formula seeing names, lost references.
 */
import { one, sheet, type CellSpec, type SpillCase } from './cases.js'
import type { NameInput } from '../../../src/contract.js'

/** Data on S: A1:A5 1…5, B1:B5 10…50, C1:C3 100…300. */
const DATA: Record<string, CellSpec> = { A1: 1, A2: 2, A3: 3, A4: 4, A5: 5, B1: 10, B2: 20, B3: 30, B4: 40, B5: 50, C1: 100, C2: 200, C3: 300 }

const named = (id: string, about: string, names: NameInput[], cells: Record<string, CellSpec>, ranges: string[], more: Partial<SpillCase> = {}): SpillCase =>
  one(id, about, { ...DATA, ...cells }, ranges, { book: { sheets: more.book?.sheets ?? [], names }, ...more })

export function buildSpillCases(): SpillCase[] {
  return [
    named(
      'constants',
      'constants: a number, text, an array of a row and of a column, a boolean, an error',
      [
        { name: 'rate', ref: '0.2' },
        { name: 'txt', ref: '"abc"' },
        { name: 'arr', ref: '{1,2,3}' },
        { name: 'col', ref: '{1;2}' },
        { name: 'flag', ref: 'TRUE' },
        { name: 'err', ref: '#N/A' },
        { name: 'neg', ref: '-5' },
      ],
      { E1: '=rate*100', E2: '=rate', E3: '=txt&"x"', E4: '=SUM(arr)', E5: '=INDEX(arr,2)', E6: '=COLUMNS(arr)', E7: '=flag', E8: '=ISNA(err)', E9: '=neg*2', E10: '=arr', E11: '=col', E13: '=TYPE(arr)', E14: '=ROWS(col)', E15: '=arr*2' },
      ['E1:G16'],
    ),
    named(
      'formulas',
      'formulas: over a cell, a whole column, an array, a reference from INDEX, another name',
      [
        { name: 'dbl', ref: "'S'!$A$2*2" },
        { name: 'tot', ref: "SUM('S'!$B:$B)" },
        { name: 'sq', ref: "'S'!$A$1:$A$3*10" },
        { name: 'ix', ref: "INDEX('S'!$A$1:$B$5,2,0)" },
        { name: 'nm2', ref: 'dbl+1' },
        { name: 'cnt', ref: "COUNT('S'!$A:$A)" },
      ],
      { E1: '=dbl', E2: '=tot', E3: '=SUM(sq)', E4: '=sq', E7: '=SUM(ix)', E8: '=ROWS(ix)', E9: '=COLUMNS(ix)', E10: '=ix', E11: '=SUMIF(ix,">1")', E12: '=COUNT(ix)', E13: '=nm2', E14: '=cnt', E15: '=SUMIF(sq,">10")', E16: '=COUNTIF(sq,">10")', E17: '=INDEX(ix,1,2)', E18: '=ix*2' },
      ['E1:G19'],
    ),
    named(
      'union',
      'a union of two ranges as a name',
      [{ name: 'u', ref: "'S'!$A$1:$A$3,'S'!$C$1:$C$3" }],
      { E1: '=SUM(u)', E2: '=COUNT(u)', E3: '=MAX(u)', E4: '=ROWS(u)', E5: '=INDEX(u,2,1,2)', E6: '=INDEX(u,1,1,3)', E7: '=u', E8: '=u+1', E9: '=AREAS(u)', E10: '=INDEX(u,2)', E11: '=SUMIF(u,">1")', E12: '=LARGE(u,2)' },
      ['E1:E12'],
    ),
    named(
      'relative',
      'references without $ relative to A1: moved by the cell using the name, wrapping round the sheet',
      [
        { name: 'rel', ref: "'S'!XFB1" },
        { name: 'relr', ref: "'S'!XFB1:XFB2" },
        { name: 'mix', ref: "'S'!$A1" },
        { name: 'up', ref: "'S'!$A1048576" },
        { name: 'relf', ref: "'S'!XFB1*10" },
        { name: 'r', ref: 'ROW()' },
        { name: 'cl', ref: 'COLUMN()' },
      ],
      { D1: '=rel', D3: '=rel', D2: '=SUM(relr)', E4: '=mix', E5: '=up', D5: '=relf', F7: '=r', F8: '=cl', D10: '=relr*2' },
      ['D1:F11', 'T!D2:D3'],
      { book: { sheets: [sheet('T', { D2: '=rel', D3: '=r' })] } },
    ),
    named(
      'scope',
      'a local name before the global one, a name qualified by its sheet, the names a name sees',
      [
        { name: 'x', ref: '1' },
        { name: 'x', ref: '2', scope: 'T' },
        { name: 'y', ref: '10' },
        { name: 'y', ref: '20', scope: 'T' },
        { name: 'g', ref: 'y+0' },
        { name: 'L', ref: 'y*2', scope: 'T' },
      ],
      { E1: '=x', E2: '=T!x', E4: '=T!L', E5: '=T!y', E6: '=g' },
      ['E1:E6', 'T!A1:A4'],
      { book: { sheets: [sheet('T', { A1: '=x', A2: '=g', A3: '=L', A4: '=T!x' })] } },
    ),
    named('sheet-global', 'a global name qualified by a sheet without a local one', [{ name: 'x', ref: '1' }], { E1: '=S!x' }, ['E1']),
    named('lost', 'a name whose reference Excel lost', [{ name: 'lost', ref: "'S'!#REF!" }], { E1: '=lost', E2: '=IFERROR(lost,7)' }, ['E1:E2']),
    named('unknown-function', 'a name over a function Excel does not have', [{ name: 'bad', ref: 'FOO(1)' }], { E1: '=bad', E2: '=ISERROR(bad)' }, ['E1:E2']),
    named(
      'chain',
      'names over names, a whole column, LET in a name and a LET name over a defined one',
      [
        { name: 'a', ref: 'b*2' },
        { name: 'b', ref: "'S'!$A$1+c" },
        { name: 'c', ref: '10' },
        { name: 'wc', ref: "'S'!$A:$A" },
        { name: 'lt', ref: "LET(x,'S'!$A$2,x*3)" },
        { name: 'rate', ref: '0.5' },
      ],
      { E1: '=a', E2: '=SUM(wc)', E3: '=ROWS(wc)', E4: '=lt', E5: '=LET(rate,5,rate+rate)', E6: '=rate+LET(rate,1,rate)' },
      ['E1:E6'],
    ),
  ]
}
