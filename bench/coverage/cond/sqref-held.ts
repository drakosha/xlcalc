/**
 * Held out: the cells of a conditional format as a file other than Excel's may spell them (`sqref`): column letters in
 * lower case, `$` before letters and numbers, both, several ranges with each, whole columns, and the range of an Excel
 * 2010 rule (`<xm:sqref>`). `writeXlsx` writes each book and the spelling is put in by hand (`raw`); every cell of
 * S!A1:D6 holds a number and the rule fills the cells it covers, so the looks show which cells Excel takes. Predicted
 * before Excel ran it (`sqref-held.predicted.json`); the set `valid/sqref-held` spells data validation's ranges alike.
 */
import type { BookInput, CellInput, ConditionalFormat } from '../../../src/contract.js'
import type { CondCase, CondCases } from './cases.js'

const SHEET = 'xl/worksheets/sheet1.xml'

/** S!A1:D6 numbered, one rule over `ranges` with the formula `formula`, and the sheet L its formula may name. */
function book(ranges: string[], formula = 'TRUE'): BookInput {
  const cells: Record<string, CellInput> = {}
  for (const c of 'ABCD') for (let r = 1; r <= 6; r++) cells[`${c}${r}`] = { value: r }
  const rule: ConditionalFormat = { type: 'expression', ranges, formula, style: { fill: { color: '#FFC7CE' } } }
  return { sheets: [{ name: 'S', cells, conditionalFormats: [rule] }, { name: 'L', cells: { A1: { value: 1 } } }] }
}

/** A case whose rule over `written` is spelled `to` in the file. */
const spelled = (id: string, book: string, written: string, to: string, x14 = false): CondCase => ({
  id,
  book,
  read: [['S', 'A1:D6']],
  raw: [x14 ? { part: SHEET, from: `<xm:sqref>${written}</xm:sqref>`, to: `<xm:sqref>${to}</xm:sqref>` } : { part: SHEET, from: `sqref="${written}"`, to: `sqref="${to}"` }],
})

export function buildCondCases(): CondCases {
  return {
    books: { one: book(['A1:B2']), areas: book(['A1:A2', 'C4:D5']), column: book(['C:C']), x14: book(['A1:B2'], "'L'!$A$1>0") },
    cases: [
      spelled('lower', 'one', 'A1:B2', 'a1:b2'),
      spelled('dollar', 'one', 'A1:B2', '$A$1:$B$2'),
      spelled('mixed', 'one', 'A1:B2', '$a$1:B$2'),
      spelled('areas', 'areas', 'A1:A2 C4:D5', 'a1:A2 $C$4:d5'),
      spelled('column', 'column', 'C:C', '$c:$c'),
      spelled('x14', 'x14', 'A1:B2', '$a$1:b2', true),
    ],
  }
}
