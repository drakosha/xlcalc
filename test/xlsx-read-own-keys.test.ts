/**
 * The reader looks a file's words up in its tables by their own keys only: a word that is the name of an object's
 * method (`toString`, `constructor`) is a word the model does not know, read as any other unknown one.
 */
import { describe, expect, it } from 'vitest'
import { readXlsx } from '../src/xlsx.js'
import { codes, pack, parts, sheet, storedValues } from './xlsx-fixtures.js'

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const rels = (items: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`
const one = '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'

describe.each(['toString', 'constructor'])('xlsx reader: the word %s from a file', (word) => {
  it('a stored error is #ERROR! with an issue, as an unknown error text', async () => {
    const r = await sheet(`<row r="1"><c r="A1" t="e"><f>1/0</f><v>${word}</v></c></row>`)
    expect(storedValues(r)).toEqual({ A1: { error: 'ERROR' } })
    expect(codes(r.issues)).toContain('errorValue A1')
  })

  it('a rule of data validation of that type is left out with an issue', async () => {
    const f = parts({ sheets: [{ name: 'S', body: `${one}<dataValidations count="1"><dataValidation type="${word}" sqref="A1"><formula1>1</formula1></dataValidation></dataValidations>` }] })
    const r = await readXlsx(await pack(f))
    expect(r.book.sheets[0].validations).toBeUndefined()
    expect(r.issues.map((i) => i.message).join('\n')).toContain(`unknown type "${word}"`)
  })

  it('an icon set of that name is left out as one not in the model', async () => {
    const rule = `<cfRule type="iconSet" priority="1"><iconSet iconSet="${word}"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>`
    const f = parts({ sheets: [{ name: 'S', body: `${one}<conditionalFormatting sqref="A1:A3">${rule}</conditionalFormatting>` }] })
    const r = await readXlsx(await pack(f))
    expect(r.book.sheets[0].conditionalFormats).toBeUndefined()
    expect(r.issues.map((i) => i.message).join('\n')).toContain(`the icon set ${word} is not in the model`)
  })

  it('a part linked by a relationship of that kind is passed over, from the workbook and from a sheet', async () => {
    const f = parts({ sheets: [{ name: 'S', body: one }] })
    f['xl/_rels/workbook.xml.rels'] = f['xl/_rels/workbook.xml.rels']!.replace('</Relationships>', `<Relationship Id="rIdX" Type="${REL}/${word}" Target="x.xml"/></Relationships>`)
    f['xl/worksheets/_rels/sheet1.xml.rels'] = rels(`<Relationship Id="rIdY" Type="${REL}/${word}" Target="../y.xml"/>`)
    const r = await readXlsx(await pack(f))
    expect(r.book.sheets[0].cells).toEqual({ A1: { value: 1 } })
    expect(r.issues).toEqual([])
  })
})
