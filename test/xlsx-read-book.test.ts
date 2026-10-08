/**
 * The reader on the workbook: sheet order and kinds, hidden sheets, names of every kind, relationships with
 * absolute targets and other case, a package without content types, a wrong `dimension`, calculation settings,
 * columns, rows and merged areas, the locale property, and the parts of a sheet the model has no place for.
 */
import { describe, expect, it } from 'vitest'
import { readXlsx } from '../src/xlsx.js'
import { codes, pack, parts, read, sheet } from './xlsx-fixtures.js'

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const rels = (items: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`
const one = (v: number) => `<sheetData><row r="1"><c r="A1"><v>${v}</v></c></row></sheetData>`

describe('xlsx reader: sheets', () => {
  it('sheets in the order the workbook lists them, not by their parts; hidden and very hidden sheets are read and reported', async () => {
    const f = parts({ sheets: [{ name: 'First', body: one(1) }, { name: 'Second', body: one(2), attrs: 'state="hidden"' }, { name: 'Third', body: one(3), attrs: 'state="veryHidden"' }] })
    // the workbook lists the third part first
    f['xl/workbook.xml'] = f['xl/workbook.xml']!.replace('<sheet name="First" sheetId="1" r:id="rId1"/>', '').replace('</sheets>', '<sheet name="First" sheetId="1" r:id="rId1"/></sheets>')
    const r = await readXlsx(await pack(f))
    expect(r.book.sheets.map((s) => [s.name, s.cells.A1.value])).toEqual([['Second', 2], ['Third', 3], ['First', 1]])
    expect(r.issues.map((i) => `${i.code} ${i.sheet}`)).toEqual(['hidden Second', 'hidden Third'])
  })

  it('chart, dialog and macro sheets among the sheets are left out; a name scoped to one too', async () => {
    const f = parts({ sheets: [{ name: 'Data', body: one(1) }], workbook: { after: '<definedNames><definedName name="onChart" localSheetId="1">Data!$A$1</definedName><definedName name="onData" localSheetId="2">Data!$A$1</definedName></definedNames>' } })
    f['xl/workbook.xml'] = f['xl/workbook.xml']!.replace('<sheets>', '<sheets><sheet name="Chart1" sheetId="9" r:id="rIdC"/><sheet name="Macro1" sheetId="8" r:id="rIdM"/>')
    f['xl/_rels/workbook.xml.rels'] = f['xl/_rels/workbook.xml.rels']!.replace('</Relationships>', `<Relationship Id="rIdC" Type="${REL}/chartsheet" Target="chartsheets/sheet1.xml"/><Relationship Id="rIdM" Type="http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" Target="macrosheets/sheet1.xml"/></Relationships>`)
    const r = await readXlsx(await pack(f))
    expect(r.book.sheets.map((s) => s.name)).toEqual(['Data'])
    expect(r.book.names).toEqual([{ name: 'onData', ref: 'Data!$A$1', scope: 'Data' }])
    expect(r.issues.map((i) => `${i.code} ${i.sheet ?? i.name}`)).toEqual(['name onChart', 'sheetKind Chart1', 'sheetKind Macro1'])
  })

  it('relationships with absolute targets, other case and percent escapes; no content types part; a wrong dimension', async () => {
    const r = await readXlsx(
      await pack({
        '_rels/.rels': rels(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="/XL/Workbook.xml"/>`),
        'xl/workbook.xml': '<workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId7"/></sheets></workbook>',
        'xl/_rels/workbook.xml.rels': rels(`<Relationship Id="rId7" Type="${REL}/worksheet" Target="/xl/worksheets/My%20Sheet.XML"/><Relationship Id="rId8" Type="${REL}/sharedStrings" Target="../xl/./sharedStrings.xml"/>`),
        'xl/worksheets/my sheet.xml': '<worksheet><dimension ref="A1:A1"/><sheetData><row r="5"><c r="Z5" t="s"><v>0</v></c></row></sheetData></worksheet>',
        'xl/sharedStrings.xml': '<sst><si><t>far</t></si></sst>',
      }),
    )
    expect(r.book.sheets).toEqual([{ name: 'S', cells: { Z5: { value: 'far' } } }])
  })

  it('a Strict file: its namespaces and relationship types, dates as ISO text', async () => {
    const STRICT = 'http://purl.oclc.org/ooxml/officeDocument/relationships'
    const r = await readXlsx(
      await pack({
        '_rels/.rels': rels(`<Relationship Id="rId1" Type="${STRICT}/officeDocument" Target="xl/workbook.xml"/>`),
        'xl/workbook.xml': `<workbook xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:r="${STRICT}" conformance="strict"><workbookPr dateCompatibility="1"/><sheets><sheet name="Strict" sheetId="1" r:id="rId1"/></sheets></workbook>`,
        'xl/_rels/workbook.xml.rels': rels(`<Relationship Id="rId1" Type="${STRICT}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${STRICT}/sharedStrings" Target="sharedStrings.xml"/>`),
        'xl/sharedStrings.xml': '<sst xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><si><t>строгий</t></si></sst>',
        'xl/worksheets/sheet1.xml':
          '<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><sheetData><row r="1"><c r="A1" t="d"><v>2025-01-31T12:00:00</v></c><c r="B1" t="s"><v>0</v></c><c r="C1"><f>A1+1</f><v>45689.5</v></c></row></sheetData></worksheet>',
      }),
    )
    expect(r.book.sheets).toEqual([{ name: 'Strict', cells: { A1: { value: 45688.5 }, B1: { value: 'строгий' }, C1: { formula: 'A1+1' } } }])
    expect(r.cached[0].cells).toEqual({ C1: 45689.5 })
  })

  it('a sheet whose part is missing is left out and reported', async () => {
    const r = await read({ sheets: [{ name: 'S', body: one(1) }, { name: 'Gone', body: one(2) }], parts: { 'xl/worksheets/sheet2.xml': undefined } })
    expect(r.book.sheets.map((s) => s.name)).toEqual(['S'])
    expect(codes(r.issues)).toEqual(['part'])
  })

  it('a file that is not a zip, or has no workbook, is an error', async () => {
    await expect(readXlsx(new TextEncoder().encode('not a zip'))).rejects.toThrow(/not a readable zip/)
    await expect(readXlsx(await pack({ 'a.txt': 'x' }))).rejects.toThrow(/no workbook part/)
  })
})

describe('xlsx reader: names', () => {
  const names = (defs: string) => read({ sheets: [{ name: 'Data', body: one(1) }, { name: 'Other', body: one(2) }], workbook: { after: `<definedNames>${defs}</definedNames>` } })

  it('references, scopes by sheet index, whole columns, lost references', async () => {
    const r = await names(
      '<definedName name="rate">Data!$A$1</definedName><definedName name="local" localSheetId="1">Other!$A$1:$B$2</definedName><definedName name="col">\'Data\'!$D:$D</definedName><definedName name="lost">#REF!</definedName><definedName name="lostOn">Data!#REF!</definedName>',
    )
    expect(r.book.names).toEqual([
      { name: 'rate', ref: 'Data!$A$1' },
      { name: 'local', ref: 'Other!$A$1:$B$2', scope: 'Other' },
      { name: 'col', ref: "'Data'!$D:$D" },
      { name: 'lost', ref: '#REF!' },
      { name: 'lostOn', ref: 'Data!#REF!' },
    ])
    expect(r.issues).toEqual([])
  })

  it('constants, formulas, unions and references across sheets are names; built-in names are left out; the placeholders of newer functions are no names', async () => {
    const r = await names(
      '<definedName name="tax">0.2</definedName><definedName name="twice">Data!$A$1*2</definedName><definedName name="both">Data!$A$1:$A$3,Data!$C$1:$C$3</definedName><definedName name="cube">Data:Other!$A$1</definedName><definedName name="one">_xlfn.SINGLE(Data!$A:$A)</definedName><definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$C$9</definedName><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Data!$A$1:$A$9</definedName><definedName name="_xlfn.XLOOKUP" hidden="1">#NAME?</definedName><definedName name="secret" hidden="1">Data!$B$1</definedName><definedName name="ext">[1]Data!$A$1</definedName>',
    )
    expect(r.book.names).toEqual([
      { name: 'tax', ref: '0.2' },
      { name: 'twice', ref: 'Data!$A$1*2' },
      { name: 'both', ref: 'Data!$A$1:$A$3,Data!$C$1:$C$3' },
      { name: 'cube', ref: 'Data:Other!$A$1' },
      { name: 'one', ref: '@Data!$A:$A' },
      { name: 'secret', ref: 'Data!$B$1' },
      { name: 'ext', ref: '[1]Data!$A$1' },
    ])
    expect(r.issues.map((i) => `${i.code} ${i.name}: ${i.message.split(':')[0]}`)).toEqual([
      'name _xlnm.Print_Area: a built-in name (a print area or titles, the range of a filter) is not in the model',
      'name _xlnm._FilterDatabase: a built-in name (a print area or titles, the range of a filter) is not in the model',
      'name secret: a hidden name is read as a visible one',
      'externalLink ext: references to other workbooks are not in the engine',
    ])
  })
})

describe('xlsx reader: the workbook and the sheet around the cells', () => {
  it('calculation settings, protection, links and macros are reported', async () => {
    const r = await read({
      sheets: [{ name: 'S', body: one(1) }],
      workbook: { before: '<workbookProtection lockStructure="1"/>', after: '<calcPr iterate="1" calcMode="manual" fullCalcOnLoad="1"/><externalReferences><externalReference r:id="rIdX"/></externalReferences>' },
      parts: { 'xl/externalLinks/externalLink1.xml': '<externalLink/>' },
    })
    expect(codes(r.issues)).toEqual([
      'calculation',
      'calculation',
      'protection',
      'externalLink',
    ])
  })

  it('columns, rows and merged areas; hidden and grouped lines reported; equal neighbours in one span', async () => {
    const r = await read({
      sheets: [
        {
          name: 'S',
          body:
            '<cols><col min="1" max="1" width="20" customWidth="1"/><col min="2" max="3" width="12.5" customWidth="1" style="1"/><col min="4" max="4" width="9" style="1"/><col min="5" max="5" width="9" customWidth="1" hidden="1"/><col min="6" max="6" width="9" outlineLevel="1"/></cols>' +
            '<sheetData><row r="2" ht="30" customHeight="1"/><row r="3" ht="30" customHeight="1"/><row r="4" ht="18"/><row r="5" hidden="1"/><row r="6" s="1" customFormat="1" outlineLevel="1"/></sheetData>' +
            '<mergeCells count="3"><mergeCell ref="A1:B2"/><mergeCell ref="B2:C3"/><mergeCell ref="D4:D4"/></mergeCells>',
        },
      ],
      styles: '<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="10"/></cellXfs>',
    })
    const s = r.book.sheets[0]
    expect(s.columns).toEqual({ A: { width: 20 }, 'B:C': { width: 12.5, format: '0.00%' }, D: { format: '0.00%' }, E: { width: 0 } })
    expect(s.rows).toEqual({ '2:3': { height: 30 }, 5: { height: 0 }, 6: { format: '0.00%' } })
    expect(s.merges).toEqual(['A1:B2'])
    expect(codes(r.issues)).toEqual(['hidden 5:5 E:E', 'outline 6:6 F:F', 'part B2:C3'])
  })

  it('what a sheet has that the model does not: each reported once for the sheet', async () => {
    const body =
      '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" state="frozen"/></sheetView></sheetViews>' +
      one(1) +
      '<sheetProtection sheet="1"/><autoFilter ref="A1:A9"/><conditionalFormatting sqref="A1"><cfRule type="cellIs"/></conditionalFormatting><conditionalFormatting sqref="B1"/><dataValidations count="1"><dataValidation sqref="A1"/></dataValidations><hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks><pageSetup orientation="landscape"/><headerFooter><oddHeader>x</oddHeader></headerFooter><drawing r:id="rId2"/><tableParts count="1"><tablePart r:id="rId3"/></tableParts><extLst><ext><x14:sparklineGroups/></ext></extLst>'
    const f = parts({ sheets: [{ name: 'S', body }] })
    f['xl/worksheets/_rels/sheet1.xml.rels'] = rels(`<Relationship Id="rId4" Type="${REL}/comments" Target="../comments1.xml"/><Relationship Id="rId5" Type="${REL}/pivotTable" Target="../pivotTables/pivotTable1.xml"/>`)
    const r = await readXlsx(await pack(f))
    expect(r.issues.map((i) => i.code)).toEqual(['conditionalFormat', 'dataValidation', 'hyperlink', 'autoFilter', 'protection', 'drawing', 'sparkline', 'print', 'frozenPanes', 'comment', 'pivotTable'])
  })

  it('the locale from its custom property; an unknown one reported; the default font other than Calibri 11 reported', async () => {
    const custom = (tag: string) => ({
      '_rels/.rels': rels(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="${REL}/custom-properties" Target="docProps/custom.xml"/>`),
      'docProps/custom.xml': `<Properties xmlns:vt="v"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="xlcalc.locale"><vt:lpwstr>${tag}</vt:lpwstr></property></Properties>`,
    })
    expect((await read({ sheets: [{ name: 'S', body: one(1) }], parts: custom('ru-RU') })).book.locale).toBe('ru-RU')
    const unknown = await read({ sheets: [{ name: 'S', body: one(1) }], parts: custom('de-DE'), styles: '<fonts count="1"><font><sz val="10"/><name val="Arial"/></font></fonts>' })
    expect(unknown.book.locale).toBeUndefined()
    expect(unknown.issues.map((i) => i.message)).toEqual(["the default font Arial 10 is read as Calibri 11, the model's default", 'the locale "de-DE" is not one of the package\'s, left out'])
  })

  it('a sheet without sheetData is an empty sheet', async () => {
    const r = await sheet('')
    expect(r.book.sheets).toEqual([{ name: 'S', cells: {} }])
  })
})
