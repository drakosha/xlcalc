/**
 * A file whose sheet names Excel would not take (x-xlsx-16) reads with the names repaired and reported, and every
 * reference to such a sheet follows its new name.
 */
import { describe, expect, it } from 'vitest'
import { compute } from '../src/index.js'
import { ruRU } from '../src/locales/ru-RU.js'
import { readXlsx } from '../src/xlsx.js'
import { pack, parts, read } from './xlsx-fixtures.js'

const one = (v: number) => `<sheetData><row r="1"><c r="A1"><v>${v}</v></c></row></sheetData>`
const namesOf = (r: Awaited<ReturnType<typeof read>>) => r.book.sheets.map((s) => s.name)
const repairs = (r: Awaited<ReturnType<typeof read>>) => r.issues.filter((i) => i.code === 'part').map((i) => [i.name, i.sheet])

describe('xlsx: sheet names Excel would not take (x-xlsx-16)', () => {
  it.each([
    ['History', ['History'], ['History (2)']],
    ['a reserved name in another case', ['hISTORY'], ['hISTORY (2)']],
    ['more than 31 characters', ['a'.repeat(32)], ['a'.repeat(31)]],
    ['empty', ['Data', ''], ['Data', 'Sheet2']],
    ['a duplicate in another case', ['S', 's'], ['S', 's (2)']],
    ['an apostrophe at either end', ["'x", "y'"], ['_x', 'y_']],
    ['forbidden characters', ['a[1]*?:/\\b'], ['a_1______b']],
    ['a repaired name taken by a later sheet', ['History', 'History (2)'], ['History (3)', 'History (2)']],
    ['a long duplicate numbered within 31 characters', ['b'.repeat(35), 'b'.repeat(33)], ['b'.repeat(31), `${'b'.repeat(27)} (2)`]],
    ['a surrogate pair at the cut', [`${'c'.repeat(30)}😀z`], ['c'.repeat(30)]],
  ])('%s: the sheet is read under a name Excel takes, with the file\'s name reported', async (_, given, names) => {
    const r = await read({ sheets: given.map((name, i) => ({ name, body: one(i) })) })
    expect(namesOf(r)).toEqual(names)
    expect(r.cached.sheets.map((s) => s.name)).toEqual(names)
    expect(repairs(r)).toEqual(given.flatMap((g, i) => (g === names[i] ? [] : [[g, names[i]]])))
  })

  it('references to the sheet in formulas, shared formulas, names, calculated columns, rules and validations follow its new name', async () => {
    const data = '<sheetData><row r="1"><c r="A1"><v>2</v></c><c r="B1"><v>3</v></c></row><row r="2"><c r="A2"><v>5</v></c></row></sheetData>'
    const rows =
      '<row r="1"><c r="A1"><f>History!A1*10</f></c><c r="B1"><f>\'History\'!B1+history!A1</f></c><c r="C1"><f>SUM(History:Other!A1)</f></c><c r="D1"><f>N</f></c><c r="E1"><f>Other!A1</f></c></row>' +
      '<row r="2"><c r="A2"><f t="shared" ref="A2:A3" si="0">History!A1+1</f></c></row><row r="3"><c r="A3"><f t="shared" si="0"/></c></row>' +
      '<row r="5"><c r="A5" t="s"><v>0</v></c></row><row r="6"><c r="A6"><v>1</v></c></row>'
    const other =
      `<sheetData>${rows}</sheetData>` +
      '<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1" dxfId="0"><formula>History!$A$1&gt;1</formula></cfRule></conditionalFormatting>' +
      '<dataValidations count="1"><dataValidation type="list" sqref="F1"><formula1>History!$A$1:$A$2</formula1></dataValidation></dataValidations>' +
      '<tableParts count="1"><tablePart r:id="rT"/></tableParts>'
    const f = parts({
      sheets: [{ name: 'History', body: data }, { name: 'Other', body: other }],
      strings: ['Qty'],
      styles: '<dxfs count="1"><dxf><font><b/></font></dxf></dxfs>',
      workbook: { after: '<definedNames><definedName name="N">History!$B$1</definedName><definedName name="Own" localSheetId="0">History!$A$2</definedName></definedNames>' },
    })
    f['xl/worksheets/_rels/sheet2.xml.rels'] = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>'
    f['xl/tables/table1.xml'] = '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="A5:A6"><tableColumns count="1"><tableColumn id="1" name="Qty"><calculatedColumnFormula>History!$A$1</calculatedColumnFormula></tableColumn></tableColumns></table>'
    const r = await readXlsx(await pack(f))
    expect(namesOf(r)).toEqual(['History (2)', 'Other'])
    const cells = r.book.sheets[1].cells
    expect([cells.A1, cells.B1, cells.C1, cells.E1, cells.A2, cells.A3].map((c) => c.formula)).toEqual([
      "'History (2)'!A1*10",
      "'History (2)'!B1+'History (2)'!A1",
      "SUM('History (2):Other'!A1)",
      'Other!A1',
      "'History (2)'!A1+1",
      "'History (2)'!A2+1",
    ])
    expect(r.book.names).toEqual([{ name: 'N', ref: "'History (2)'!$B$1" }, { name: 'Own', ref: "'History (2)'!$A$2", scope: 'History (2)' }])
    expect(r.book.sheets[1].conditionalFormats?.[0]).toMatchObject({ formula: "'History (2)'!$A$1>1" })
    expect(r.book.sheets[1].validations?.[0]).toMatchObject({ formula1: "'History (2)'!$A$1:$A$2" })
    expect(r.book.sheets[1].tables?.[0].columns[0].formula).toBe("'History (2)'!$A$1")
    expect(r.issues.map((i) => [i.code, i.sheet, i.name])).toEqual([['part', 'History (2)', 'History']])
    for (const execution of ['closures', 'eval'] as const) {
      const out = compute(r.book, { execution }).sheets[1].cells
      expect([out.A1, out.B1, out.C1, out.D1, out.A2, out.A3].map((c) => c.value)).toEqual([20, 5, 22, 3, 3, 6])
    }
  })

  it('a hyperlink\'s location on the sheet follows its new name, as a formula\'s reference does (link-repair)', async () => {
    const locations = ['History!A1', "'history'!B2:C3", "'History'!N", 'History!A1:History!B2', 'History!#REF!', 'Other!A1', 'A1', 'N']
    const rows = locations.map((_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>${i}</v></c></row>`).join('')
    const links = locations.map((l, i) => `<hyperlink ref="A${i + 1}" location="${l}"/>`).join('')
    const r = await read({ sheets: [{ name: 'History', body: one(1) }, { name: 'Other', body: `<sheetData>${rows}</sheetData><hyperlinks>${links}</hyperlinks>` }] })
    expect(Object.values(r.book.sheets[1].cells).map((c) => c.link?.location)).toEqual([
      "'History (2)'!A1",
      "'History (2)'!B2:C3",
      "'History (2)'!N",
      "'History (2)'!A1:'History (2)'!B2",
      "'History (2)'!#REF!",
      'Other!A1',
      'A1',
      'N',
    ])
  })

  it('a reference to a duplicate name goes to the first sheet with it, which keeps its name', async () => {
    const r = await read({ sheets: [{ name: 'S', body: one(1) }, { name: 's', body: one(2) }, { name: 'U', body: '<sheetData><row r="1"><c r="A1"><f>s!A1</f></c></row></sheetData>' }] })
    expect(namesOf(r)).toEqual(['S', 's (2)', 'U'])
    expect(r.book.sheets[2].cells.A1.formula).toBe('s!A1')
  })

  it('the issue is in the language of the read', async () => {
    const data = await pack(parts({ sheets: [{ name: 'History', body: one(1) }] }))
    expect((await readXlsx(data)).issues.map((i) => i.message)).toEqual(['the sheet name "History" is not one Excel takes: the sheet is read as "History (2)"'])
    expect((await readXlsx(data, { locale: ruRU })).issues.map((i) => i.message)).toEqual(['Имя листа «History» Excel не принимает: лист прочитан как «History (2)»'])
  })
})
