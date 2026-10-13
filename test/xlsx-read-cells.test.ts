/**
 * The reader on cells: placement by address and by position, value types, text with Excel's escapes, rich and
 * phonetic runs, errors, dates of the 1904 system and of Strict files, cell metadata. The first cases are the
 * an earlier reader's regressions: a self-closing cell between values, cells without `r`, prefixed tags.
 */
import { describe, expect, it } from 'vitest'
import { cellsOf } from './looks-book.js'
import { codes, read, sheet, storedValues } from './xlsx-fixtures.js'

const DATE_STYLES = '<numFmts count="2"><numFmt numFmtId="164" formatCode="0 m"/><numFmt numFmtId="165" formatCode="dd.mm.yyyy"/></numFmts><cellXfs count="4"><xf numFmtId="0"/><xf numFmtId="14"/><xf numFmtId="164"/><xf numFmtId="165"/></cellXfs>'

describe('xlsx reader: cells', () => {
  it('places cells by address: an omitted cell stays empty, columns do not shift', async () => {
    const r = await sheet('<row r="1"><c r="A1" t="s"><v>0</v></c><c r="C1" t="s"><v>1</v></c></row><row r="3"><c r="B3"><v>42</v></c></row>', { strings: ['Имя', 'Город'] })
    expect(r.book.sheets[0].cells).toEqual({ A1: { value: 'Имя' }, C1: { value: 'Город' }, B3: { value: 42 } })
  })

  it('a self-closing cell between values does not swallow the next one', async () => {
    const r = await sheet('<row r="1"><c r="A1"><v>10</v></c><c r="B1" s="0"/><c r="C1" t="s"><v>0</v></c><c r="D1"/><c r="E1"><v>50</v></c></row>', { strings: ['DATA_C'] })
    expect(r.book.sheets[0].cells).toEqual({ A1: { value: 10 }, C1: { value: 'DATA_C' }, E1: { value: 50 } })
  })

  it('cells and rows without `r` are placed after the previous ones', async () => {
    const r = await sheet('<row><c t="s"><v>0</v></c><c><v>7</v></c><c t="s"><v>1</v></c></row><row><c r="C2"><v>1</v></c><c><v>2</v></c></row>', { strings: ['Альфа', 'Бета'] })
    expect(r.book.sheets[0].cells).toEqual({ A1: { value: 'Альфа' }, B1: { value: 7 }, C1: { value: 'Бета' }, C2: { value: 1 }, D2: { value: 2 } })
  })

  it('a cell outside the sheet is left out and reported; lower-case addresses are spelled in upper case', async () => {
    const r = await sheet('<row r="1"><c r="a1"><v>1</v></c><c r="XFE1"><v>2</v></c></row><row r="1048577"><c r="A1048577"><v>3</v></c></row>')
    expect(r.book.sheets[0].cells).toEqual({ A1: { value: 1 } })
    expect(codes(r.issues)).toEqual(['part XFE1 A1048577'])
  })

  it('elements and attributes the reader does not know are passed over', async () => {
    const r = await sheet(
      '<row r="1" spans="1:3" x14ac:dyDescent="0.25" thickBot="1"><c r="A1" ph="1" x:foo="bar"><extLst><ext uri="x"/></extLst><v>1</v></c><c r="B1" t="s"><v>0</v><extra/></c></row><row r="2"/>',
      { strings: ['s'] },
    )
    expect(r.book.sheets[0].cells).toEqual({ A1: { value: 1 }, B1: { value: 's' } })
    expect(r.issues).toEqual([])
  })

  it('tags with a namespace prefix', async () => {
    const r = await read({ sheets: [{ name: 'S', body: '<x:sheetData><x:row r="1"><x:c r="A1"><x:v>99</x:v></x:c><x:c r="B1" t="s"><x:v>0</x:v></x:c><x:c r="C1"><x:f>A1*2</x:f><x:v>198</x:v></x:c></x:row></x:sheetData>' }], strings: ['нс'] })
    expect(r.book.sheets[0].cells).toEqual({ A1: { value: 99 }, B1: { value: 'нс' }, C1: { formula: 'A1*2' } })
    expect(storedValues(r)).toEqual({ C1: 198 })
  })

  it('value types: inline and formula strings, booleans, errors, numbers in exponent form', async () => {
    const r = await sheet(
      '<row r="1"><c r="A1" t="inlineStr"><is><t>Вася</t></is></c><c r="B1" t="str"><v>plain</v></c><c r="C1" t="b"><v>1</v></c><c r="D1" t="b"><v>0</v></c><c r="E1"><v>1.5E-3</v></c><c r="F1" t="n"><v></v></c></row>' +
        '<row r="2"><c r="A2" t="str"><f>"a"&amp;"b"</f><v>ab</v></c><c r="B2" t="e"><f>1/0</f><v>#DIV/0!</v></c><c r="C2" t="b"><f>TRUE</f><v>1</v></c><c r="D2" t="e"><f>NA()</f><v>#N/A</v></c></row>',
    )
    expect(r.book.sheets[0].cells).toEqual({ A1: { value: 'Вася' }, B1: { value: 'plain' }, C1: { value: true }, D1: { value: false }, E1: { value: 0.0015 }, A2: { formula: '"a"&"b"' }, B2: { formula: '1/0' }, C2: { formula: 'TRUE' }, D2: { formula: 'NA()' } })
    expect(storedValues(r)).toEqual({ A2: 'ab', B2: { error: 'DIV0' }, C2: true, D2: { error: 'NA' } })
    expect(r.issues).toEqual([])
  })

  it("text with Excel's escapes, entities, a carriage return and spaces kept", async () => {
    const r = await sheet(
      '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c><c r="D1" t="s"><v>3</v></c><c r="E1" t="inlineStr"><is><t xml:space="preserve">  two  </t></is></c></row>',
      { strings: ['_x005F_x0041_ is literal', 'vertical_x000B_tab', 'cr_x000D_\r\nlf', '&lt;&amp;&gt;&quot;&#x41;&#66;'] },
    )
    expect(r.book.sheets[0].cells).toEqual({ A1: { value: '_x0041_ is literal' }, B1: { value: 'vertical\u000btab' }, C1: { value: 'cr\r\nlf' }, D1: { value: '<&>"AB' }, E1: { value: '  two  ' } })
  })

  it('rich text is read as its plain text and reported; phonetic runs are not part of the text', async () => {
    const rich = '<r><rPr><b/></rPr><t>bold</t></r><r><t xml:space="preserve"> plain</t></r>'
    const phonetic = '<r><t>東京</t></r><rPh sb="0" eb="2"><t>トウキョウ</t></rPh><phoneticPr fontId="1"/>'
    const r = await sheet('<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="inlineStr"><is><t>no runs</t><rPh><t>x</t></rPh></is></c><c r="D1" t="inlineStr"><is><r><t>inline</t></r></is></c></row>', { strings: [rich, phonetic] })
    expect(r.book.sheets[0].cells).toEqual({ A1: { value: 'bold plain' }, B1: { value: '東京' }, C1: { value: 'no runs' }, D1: { value: 'inline' } })
    expect(codes(r.issues)).toEqual(['richText A1 B1 D1'])
  })

  it('an error typed into a cell is read as a formula of that error; an error the model has no code for is reported', async () => {
    const r = await sheet('<row r="1"><c r="A1" t="e"><v>#N/A</v></c><c r="B1" t="e"><v>#NULL!</v></c><c r="C1" t="e"><v>#GETTING_DATA</v></c><c r="D1" t="e"><f>X()</f><v>#BLOCKED!</v></c></row>')
    expect(r.book.sheets[0].cells).toEqual({ A1: { formula: '#N/A' }, B1: { formula: '#NULL!' }, D1: { formula: 'X()' } })
    expect(storedValues(r)).toEqual({ A1: { error: 'NA' }, B1: { error: 'NULL' }, D1: { error: 'ERROR' } })
    expect(codes(r.issues)).toEqual(['errorValue C1', 'errorValue D1', 'errorConstant A1 B1'])
  })

  it('numbers stay numbers whatever their format; a built-in format is read by its number', async () => {
    const r = await sheet('<row r="1"><c r="A1" s="1"><v>45000</v></c><c r="B1" s="2"><v>5</v></c><c r="C1" s="3"><v>45000.5</v></c></row>', { styles: DATE_STYLES })
    expect(r.book.sheets[0].cells).toEqual({ A1: { value: 45000, numberFormat: 14 }, B1: { value: 5, numberFormat: '0 m' }, C1: { value: 45000.5, numberFormat: 'dd.mm.yyyy' } })
    expect(r.issues).toEqual([])
  })

  it("the 1904 date system: the book's date system, numbers as the file stores them, ISO dates counted from 1904", async () => {
    const r = await read({
      sheets: [{ name: 'S', body: '<sheetData><row r="1"><c r="A1" s="1"><v>0</v></c><c r="B1" s="3"><v>45000</v></c><c r="C1" s="2"><v>5</v></c><c r="D1"><v>7</v></c><c r="E1" s="1"><f>A1+1</f><v>1</v></c><c r="F1" t="d"><v>1904-01-02T12:00:00</v></c></row></sheetData>' }],
      styles: DATE_STYLES,
      workbook: { before: '<workbookPr date1904="1"/>' },
    })
    expect(r.book.dateSystem).toBe(1904)
    expect(r.book.sheets[0].cells).toEqual({ A1: { value: 0, numberFormat: 14 }, B1: { value: 45000, numberFormat: 'dd.mm.yyyy' }, C1: { value: 5, numberFormat: '0 m' }, D1: { value: 7 }, E1: { formula: 'A1+1', numberFormat: 14 }, F1: { value: 1.5 } })
    expect(storedValues(r)).toEqual({ E1: 1 })
    expect(r.issues).toEqual([])
    expect((await sheet('<row r="1"><c r="A1"><v>1</v></c></row>')).book.dateSystem).toBeUndefined()
  })

  it('ISO dates of Strict files (t="d") as serial numbers', async () => {
    const r = await sheet('<row r="1"><c r="A1" t="d"><v>2025-01-31T00:00:00</v></c><c r="B1" t="d"><v>2025-01-31T18:00:00</v></c></row>')
    expect(r.book.sheets[0].cells).toEqual({ A1: { value: 45688 }, B1: { value: 45688.75 } })
  })

  it("a cell's value metadata (vm) and a what-if data table are reported, their values kept", async () => {
    const r = await sheet('<row r="1"><c r="A1" t="e" vm="1"><v>#VALUE!</v></c><c r="B1"><f t="dataTable" ref="B1:B3" dt2D="0" dtr="0" r1="A1"/><v>3</v></c></row>')
    expect(r.book.sheets[0].cells).toEqual({ A1: { formula: '#VALUE!' }, B1: { value: 3 } })
    expect(codes(r.issues)).toEqual(['cellMetadata A1', 'dataTable B1', 'errorConstant A1'])
  })

  it('an empty cell with a style keeps its look; one with the default look is no cell', async () => {
    const r = await sheet('<row r="1"><c r="A1" s="1"/><c r="B1" s="0"/><c r="C1"/></row>', { styles: '<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="2"/></cellXfs>' })
    expect(cellsOf(r.book.sheets[0])).toEqual({ A1: { numberFormat: 2 } })
  })
})
