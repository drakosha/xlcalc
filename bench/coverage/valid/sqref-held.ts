/**
 * Held out: the cells of a rule of data validation as a file other than Excel's may spell them (`sqref`): column letters
 * in lower case, `$` before letters and numbers, both, several ranges with each, and the range of an Excel 2010 rule
 * (`<xm:sqref>` in the sheet's extension, where Excel writes a rule over another sheet). `writeXlsx` writes each book and
 * the spelling is put in by hand (`raw`); every cell of S!A1:D6 holds a number, so the cells read with a rule show which
 * cells Excel takes. Predicted before Excel ran it (`sqref-held.predicted.json`); the set `cond/sqref-held` spells
 * conditional formats' ranges alike.
 */
import type { BookInput, CellInput, DataValidation } from '../../../src/contract.js'
import type { RawEdit } from '../xlsx/sent.js'
import type { ValidCase, ValidCases } from './cases.js'

const SHEET = 'xl/worksheets/sheet1.xml'

/** S!A1:D6 numbered, one rule over `range`, and the list on sheet L a rule may take. */
function book(range: string, rule: Omit<DataValidation, 'range'> = { type: 'whole', operator: 'greaterThan', formula1: '2' }): BookInput {
  const cells: Record<string, CellInput> = {}
  for (const c of 'ABCD') for (let r = 1; r <= 6; r++) cells[`${c}${r}`] = { value: r }
  return { sheets: [{ name: 'S', cells, validations: [{ range, ...rule }] }, { name: 'L', cells: { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 } } }] }
}

/** The list rule over A1:B2 as `writeXlsx` writes it, and as Excel writes a rule over another sheet: in the sheet's extension. */
const PLAIN = `<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1:B2"><formula1>'L'!$A$1:$A$3</formula1></dataValidation></dataValidations>`
const extension = (sqref: string) =>
  '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">' +
  `<x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1"><x14:formula1><xm:f>'L'!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>${sqref}</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>`

/** A case whose rule over `written` is spelled `to` in the file. */
const spelled = (id: string, written: string, to: string): ValidCase => ({ id, book: written, raw: [{ part: SHEET, from: `sqref="${written}"`, to: `sqref="${to}"` }] })

export function buildValidCases(): ValidCases {
  const x14: RawEdit[] = [
    { part: SHEET, from: PLAIN, to: '' },
    { part: SHEET, from: '</worksheet>', to: `${extension('$a$1:b2')}</worksheet>` },
  ]
  return {
    books: { 'A1:B2': book('A1:B2'), 'A1:A2 C4:D5': book('A1:A2 C4:D5'), list: book('A1:B2', { type: 'list', formula1: "'L'!$A$1:$A$3" }) },
    cases: [
      spelled('lower', 'A1:B2', 'a1:b2'),
      spelled('dollar', 'A1:B2', '$A$1:$B$2'),
      spelled('mixed', 'A1:B2', '$a$1:B$2'),
      spelled('areas', 'A1:A2 C4:D5', 'a1:A2 $C$4:d5'),
      { id: 'x14', book: 'list', raw: x14 },
    ],
  }
}
