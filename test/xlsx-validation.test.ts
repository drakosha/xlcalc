/**
 * Data validation in xlsx files: what the writer writes reads back the same, what the file cannot hold is cut or left
 * out and reported, and the reader takes each form a file holds rules in (the sheet's own, Excel 2010's extension, the
 * list of Excel 2013's), reporting what the book has no place for. Files Excel itself saved are checked in
 * `validation-excel.test.ts`.
 */
import { describe, expect, it } from 'vitest'
import type { BookInput, DataValidation } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { codes, read } from './xlsx-fixtures.js'

const roundTrip = async (book: BookInput) => {
  const w = await writeXlsx(book)
  const r = await readXlsx(w.data)
  return { book: r.book, issues: [...w.issues, ...r.issues], data: w.data }
}

/** Every type and setting, ranges of several parts, whole columns and rows, another sheet, texts with lines and markup. */
const RULES: DataValidation[] = [
  { range: 'B2:B9', type: 'list', formula1: '"Yes,No,""quoted"""', promptTitle: 'Pick', prompt: 'Line 1\nLine 2 & <b>"x"</b>' },
  { range: 'C2:C9 E2 G5:H6', type: 'whole', formula1: '$A$1', formula2: '10', errorStyle: 'warning', errorTitle: 'Whole', error: 'From A1 to 10' },
  { range: 'D2:D9', type: 'decimal', operator: 'notEqual', formula1: '0.5', allowBlank: false },
  { range: 'F1:F20', type: 'date', operator: 'greaterThan', formula1: 'DATE(2025,1,1)', errorStyle: 'information' },
  { range: 'F30:G30', type: 'time', operator: 'lessThanOrEqual', formula1: 'TIME(12,0,0)', showPrompt: false, showError: false },
  { range: 'I2', type: 'textLength', formula1: '1', formula2: 'LEN($A$2)' },
  { range: 'J2:J9', type: 'custom', formula1: 'AND(J2>0,XLOOKUP(J2,$A:$A,$A:$A)=J2)' },
  { range: 'K2:K3', type: 'list', formula1: "'T T'!$A$1:$A$3", dropdown: false },
  { range: 'L2', type: 'list', formula1: 'Fruits' },
  { range: 'M2', type: 'any', prompt: 'just a note' },
]

/** Whole columns, and whole rows on a sheet of their own. */
const WHOLE: DataValidation[] = [{ range: 'F:F', type: 'list', formula1: '"a,b"' }]
const ROWS: DataValidation[] = [{ range: '20:21', type: 'any', prompt: 'rows' }]

const book = (): BookInput => ({
  sheets: [
    { name: 'S', cells: { A1: { value: 1 }, A2: { value: 'abc' } }, validations: RULES },
    { name: 'T T', cells: { A1: { value: 'a' } }, validations: WHOLE },
    { name: 'U', cells: {}, validations: ROWS },
  ],
  names: [{ name: 'Fruits', ref: "'T T'!$A$1:$A$3" }],
})

describe('data validation written and read back', () => {
  it('every type and setting comes back as written', async () => {
    const r = await roundTrip(book())
    expect(r.issues).toEqual([])
    expect(r.book.sheets.map((s) => s.validations)).toEqual([RULES, WHOLE, ROWS])
    const sheet = (await openArchive(r.data).text('xl/worksheets/sheet1.xml'))!
    // xlsx's defaults are false: the model's, true, are written; whole columns and rows as cells; the new function prefixed
    expect(sheet).toContain('<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="Pick" prompt="Line 1&#10;Line 2 &amp; &lt;b&gt;&quot;x&quot;&lt;/b&gt;" sqref="B2:B9">')
    const second = (await openArchive(r.data).text('xl/worksheets/sheet2.xml'))!
    expect(second).toContain('sqref="F1:F1048576"')
    expect((await openArchive(r.data).text('xl/worksheets/sheet3.xml'))!).toContain('sqref="A20:XFD21"')
    expect(sheet).toContain('<formula1>AND(J2&gt;0,_xlfn.XLOOKUP(J2,$A:$A,$A:$A)=J2)</formula1>')
    expect(sheet.indexOf('<dataValidations')).toBeGreaterThan(sheet.indexOf('</sheetData>'))
  })

  it("texts over Excel's limits are cut, a rule whose formula the file cannot hold is left out, both reported", async () => {
    const long = (n: number) => 'x'.repeat(n)
    const b: BookInput = {
      sheets: [
        {
          name: 'S',
          cells: {},
          validations: [
            { range: 'A1', type: 'any', promptTitle: long(40), prompt: long(300), errorTitle: long(33), error: long(226) },
            { range: 'B1', type: 'list', formula1: "'Nowhere'!$A$1:$A$2" },
            { range: 'C1', type: 'custom', formula1: 'SUM(' },
          ],
        },
      ],
    }
    const w = await writeXlsx(b)
    expect(codes(w.issues)).toEqual(['textLength A1', 'textLength A1', 'textLength A1', 'textLength A1', 'dataValidation B1 C1'])
    const r = await readXlsx(w.data)
    expect(r.book.sheets[0].validations).toEqual([{ range: 'A1', type: 'any', promptTitle: long(32), prompt: long(255), errorTitle: long(32), error: long(225) }])
  })
})

describe('data validation read from each form a file holds it in', () => {
  const sheet = (tail: string) => read({ sheets: [{ name: 'S', body: `<sheetData/>${tail}` }, { name: 'L', body: '<sheetData/>' }] })

  it("xlsx's defaults spelled out; types and operators that do not go together dropped; Excel's escapes and entities decoded", async () => {
    const r = await sheet(
      '<dataValidations count="3"><dataValidation type="list" operator="equal" showDropDown="1" sqref="a1:a3"><formula1>"x,y"</formula1></dataValidation>' +
        '<dataValidation type="whole" operator="between" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorStyle="stop" prompt="a_x000a_b &amp; c" sqref="B1"><formula1>1</formula1><formula2>2</formula2></dataValidation>' +
        '<dataValidation showDropDown="1" sqref="C1"/></dataValidations>',
    )
    expect(r.issues).toEqual([])
    expect(r.book.sheets[0].validations).toEqual([
      { range: 'A1:A3', type: 'list', formula1: '"x,y"', allowBlank: false, dropdown: false, showPrompt: false, showError: false },
      { range: 'B1', type: 'whole', formula1: '1', formula2: '2', prompt: 'a\nb & c' },
      { range: 'C1', type: 'any', allowBlank: false, showPrompt: false, showError: false },
    ])
  })

  it("Excel 2010's extension: rules over other sheets, beside the sheet's own", async () => {
    const r = await sheet(
      '<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1"><formula1>"a,b"</formula1></dataValidation></dataValidations>' +
        '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1"><x14:formula1><xm:f>L!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>B2:B5 D1</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>',
    )
    expect(r.issues).toEqual([])
    expect(r.book.sheets[0].validations).toEqual([
      { range: 'A1', type: 'list', formula1: '"a,b"' },
      { range: 'B2:B5 D1', type: 'list', formula1: 'L!$A$1:$A$3' },
    ])
  })

  it('what the book has no place for is reported: rules without a range or a formula, cells two rules share, IME modes, items holding commas, automatic outline styles', async () => {
    const r = await read({
      sheets: [
        {
          name: 'S',
          body:
            '<sheetPr><outlinePr applyStyles="1" summaryBelow="0"/></sheetPr><sheetData/><dataValidations count="7">' +
            '<dataValidation type="whole" sqref="A1"><formula1>1</formula1></dataValidation>' +
            '<dataValidation type="list" sqref="A2"/>' +
            '<dataValidation type="list" sqref="ZZZZ1"><formula1>"a"</formula1></dataValidation>' +
            '<dataValidation type="custom" sqref="C2:C5"><formula1>C2&gt;0</formula1></dataValidation>' +
            '<dataValidation type="custom" sqref="C1:C3 E1"><formula1>C1&gt;1</formula1></dataValidation>' +
            '<dataValidation type="list" imeMode="hiragana" sqref="F1"><formula1>"a"</formula1></dataValidation>' +
            '<dataValidation type="list" sqref="G1"><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x12ac"><x12ac:list>"p,q",r</x12ac:list></mc:Choice><mc:Fallback><formula1>"p,q,r"</formula1></mc:Fallback></mc:AlternateContent></dataValidation>' +
            '<dataValidation type="bogus" sqref="H1"><formula1>1</formula1></dataValidation></dataValidations>',
        },
      ],
    })
    expect(codes(r.issues)).toEqual(['dataValidation G1', 'dataValidation A1', 'dataValidation A2', 'dataValidation ZZZZ1', 'dataValidation C1:C3 E1', 'dataValidation F1', 'dataValidation H1', 'outline'])
    // the cells a later rule shares go to the earlier one; the later's formulas are written for its new first cell
    expect(r.book.sheets[0].validations!.map((v) => `${v.range} ${v.formula1}`)).toEqual(['C2:C5 C2>0', 'C1 E1 C1>1', 'F1 "a"', 'G1 "p,q,r"'])
    expect(r.book.sheets[0].outline).toEqual({ summaryBelow: false })
  })
})
