/**
 * The reader on the workbook: sheet order and kinds, hidden sheets, names of every kind, relationships with
 * absolute targets and other case, a package without content types, a wrong `dimension`, calculation settings,
 * columns, rows and merged areas, the locale property, and the parts of a sheet the model has no place for.
 */
import { describe, expect, it } from 'vitest'
import { readXlsx } from '../src/xlsx.js'
import { codes, pack, parts, read, rels, sheet, storedValues } from './xlsx-fixtures.js'

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const one = (v: number) => `<sheetData><row r="1"><c r="A1"><v>${v}</v></c></row></sheetData>`

describe('xlsx reader: sheets', () => {
  it('sheets in the order the workbook lists them, not by their parts; hidden and very hidden sheets are read so', async () => {
    const f = parts({ sheets: [{ name: 'First', body: one(1) }, { name: 'Second', body: one(2), attrs: 'state="hidden"' }, { name: 'Third', body: one(3), attrs: 'state="veryHidden"' }, { name: 'Fourth', body: one(4), attrs: 'state="visible"' }] })
    // the workbook lists the third part first
    f['xl/workbook.xml'] = f['xl/workbook.xml']!.replace('<sheet name="First" sheetId="1" r:id="rId1"/>', '').replace('</sheets>', '<sheet name="First" sheetId="1" r:id="rId1"/></sheets>')
    const r = await readXlsx(await pack(f))
    expect(r.book.sheets.map((s) => [s.name, s.cells.A1.value, s.hidden])).toEqual([['Second', 2, true], ['Third', 3, 'very'], ['Fourth', 4, undefined], ['First', 1, undefined]])
    expect(r.issues).toEqual([])
  })

  it('a file whose every worksheet is hidden (beside a chart sheet) shows its first worksheet, reported', async () => {
    const f = parts({ sheets: [{ name: 'Data', body: one(1), attrs: 'state="hidden"' }, { name: 'More', body: one(2), attrs: 'state="veryHidden"' }] })
    f['xl/workbook.xml'] = f['xl/workbook.xml']!.replace('<sheets>', '<sheets><sheet name="Chart1" sheetId="9" r:id="rIdC"/>')
    f['xl/_rels/workbook.xml.rels'] = f['xl/_rels/workbook.xml.rels']!.replace('</Relationships>', `<Relationship Id="rIdC" Type="${REL}/chartsheet" Target="chartsheets/sheet1.xml"/></Relationships>`)
    const r = await readXlsx(await pack(f))
    expect(r.book.sheets.map((s) => [s.name, s.hidden])).toEqual([['Data', undefined], ['More', 'very']])
    expect(r.issues.map((i) => `${i.code} ${i.sheet}`)).toEqual(['sheetKind Chart1', 'hidden Data'])
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
    expect(storedValues(r)).toEqual({ C1: 45689.5 })
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

  it('constants, formulas, unions and references across sheets are names; a hidden name is read hidden; built-in names are left out; the placeholders of newer functions are no names', async () => {
    const r = await names(
      '<definedName name="tax">0.2</definedName><definedName name="twice">Data!$A$1*2</definedName><definedName name="both">Data!$A$1:$A$3,Data!$C$1:$C$3</definedName><definedName name="cube">Data:Other!$A$1</definedName><definedName name="one">_xlfn.SINGLE(Data!$A:$A)</definedName><definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$C$9</definedName><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Data!$A$1:$A$9</definedName><definedName name="_xlfn.XLOOKUP" hidden="1">#NAME?</definedName><definedName name="secret" hidden="1">Data!$B$1</definedName><definedName name="ext">[1]Data!$A$1</definedName>',
    )
    expect(r.book.names).toEqual([
      { name: 'tax', ref: '0.2' },
      { name: 'twice', ref: 'Data!$A$1*2' },
      { name: 'both', ref: 'Data!$A$1:$A$3,Data!$C$1:$C$3' },
      { name: 'cube', ref: 'Data:Other!$A$1' },
      { name: 'one', ref: '@Data!$A:$A' },
      { name: 'secret', ref: 'Data!$B$1', hidden: true },
      { name: 'ext', ref: '[1]Data!$A$1' },
    ])
    expect(r.issues.map((i) => `${i.code} ${i.name}: ${i.message.split(':')[0]}`)).toEqual([
      'name _xlnm.Print_Area: a built-in name (a print area, print titles, a filter range) is not in the model',
      'name _xlnm._FilterDatabase: a built-in name (a print area, print titles, a filter range) is not in the model',
      'externalLink ext: a reference to another workbook that is not one of the links the file holds, or that the engine does not parse',
    ])
  })
})

describe('xlsx reader: the workbook and the sheet around the cells', () => {
  it('calculation settings the book does not keep, protection, links and macros are reported', async () => {
    const r = await read({
      sheets: [{ name: 'S', body: one(1) }],
      workbook: { before: '<workbookProtection lockStructure="1"/>', after: '<calcPr iterate="1" calcMode="manual" fullCalcOnLoad="1"/><externalReferences><externalReference r:id="rIdX"/></externalReferences>' },
      parts: { 'xl/externalLinks/externalLink1.xml': '<externalLink/>' },
    })
    expect(r.book.iteration).toEqual({})
    expect(codes(r.issues)).toEqual(['calculation', 'protection', 'externalLink'])
  })

  it('iterative calculation with its count and change; settings out of range read as the defaults and reported; precision as displayed reported', async () => {
    const calc = async (attrs: string) => {
      const r = await read({ sheets: [{ name: 'S', body: one(1) }], workbook: { after: `<calcPr calcId="191029" ${attrs}/>` } })
      return { iteration: r.book.iteration, issues: r.issues.map((i) => i.message) }
    }
    expect(await calc('iterate="1" iterateCount="7" iterateDelta="0.5"')).toEqual({ iteration: { maxIterations: 7, maxChange: 0.5 }, issues: [] })
    expect(await calc('iterate="true" iterateDelta="0"')).toEqual({ iteration: { maxChange: 0 }, issues: [] })
    expect(await calc('iterateCount="7" iterateDelta="0.5"')).toEqual({ iteration: undefined, issues: [] })
    expect(await calc('iterate="1" iterateCount="0" iterateDelta="-1"')).toEqual({
      iteration: {},
      issues: ['the maximum number of iterations "0" is not between 1 and 32,767: the default, 100, is used', 'the maximum change of iterations "-1" is not a number of 0 or more: the default, 0.001, is used'],
    })
    expect(await calc('calcMode="autoNoTable" fullPrecision="0"')).toEqual({
      iteration: undefined,
      issues: ['the calculation mode "automatic except for data tables" is not kept', 'the engine does not support precision as displayed: values keep their full precision'],
    })
  })

  it('columns, rows and merged areas; hidden lines with their sizes, a line of size 0 hidden, a height without customHeight one Excel fitted, the highest of Excel for Mac, 409.6, read as 409.5; grouped lines with their levels; equal neighbours in one span', async () => {
    const r = await read({
      sheets: [
        {
          name: 'S',
          body:
            '<cols><col min="1" max="1" width="20" customWidth="1"/><col min="2" max="3" width="12.5" customWidth="1" style="1"/><col min="4" max="4" width="9" style="1"/><col min="5" max="5" width="9" customWidth="1" hidden="1"/><col min="6" max="6" width="9" outlineLevel="1"/><col min="7" max="7" width="9" hidden="1"/><col min="8" max="8" width="0" customWidth="1"/></cols>' +
            '<sheetData><row r="2" ht="30" customHeight="1"/><row r="3" ht="30" customHeight="1"/><row r="4" ht="18"/><row r="5" hidden="1"/><row r="6" s="1" customFormat="1" outlineLevel="1"/><row r="7" ht="25" hidden="1" customHeight="1"/><row r="8" ht="0" customHeight="1"/><row r="9" ht="409.6"/></sheetData>' +
            '<mergeCells count="3"><mergeCell ref="A1:B2"/><mergeCell ref="B2:C3"/><mergeCell ref="D4:D4"/></mergeCells>',
        },
      ],
      styles: '<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="10"/></cellXfs>',
    })
    const s = r.book.sheets[0]
    expect(s.columns).toEqual({ A: { width: 20 }, 'B:C': { width: 12.5, numberFormat: 10 }, D: { numberFormat: 10 }, E: { width: 9, hidden: true }, F: { outlineLevel: 1 }, G: { hidden: true }, H: { width: 0, hidden: true } })
    expect(s.rows).toEqual({ '2:3': { height: 30 }, 4: { height: 18, bestFit: true }, 5: { hidden: true }, 6: { numberFormat: 10, outlineLevel: 1 }, 7: { height: 25, hidden: true }, 8: { height: 0, hidden: true }, 9: { height: 409.5, bestFit: true } })
    expect(s.merges).toEqual(['A1:B2'])
    expect(codes(r.issues)).toEqual(['part B2:C3'])
  })

  it('what a sheet has that the model does not: each reported once for the sheet', async () => {
    const body =
      '<sheetViews><sheetView workbookViewId="0"><pane ySplit="600" topLeftCell="A2"/></sheetView></sheetViews>' +
      one(1) +
      '<sheetProtection sheet="1"/><autoFilter ref="A1:A9"/><conditionalFormatting sqref="A1"><cfRule type="cellIs"/></conditionalFormatting><conditionalFormatting sqref="B1"/><dataValidations count="1"><dataValidation sqref="A1"/></dataValidations><pageSetup orientation="landscape"/><headerFooter><oddHeader>x</oddHeader></headerFooter><drawing r:id="rId2"/><tableParts count="1"><tablePart r:id="rId3"/></tableParts><extLst><ext><x14:sparklineGroups/></ext></extLst>'
    const f = parts({ sheets: [{ name: 'S', body }] })
    f['xl/worksheets/_rels/sheet1.xml.rels'] = rels(`<Relationship Id="rId5" Type="${REL}/pivotTable" Target="../pivotTables/pivotTable1.xml"/>`)
    const r = await readXlsx(await pack(f))
    // conditional formats are read: a rule without its formula is left out
    expect(r.issues.map((i) => i.code)).toEqual(['protection', 'drawing', 'sparkline', 'print', 'pivotTable', 'autoFilter', 'frozenPanes', 'conditionalFormat'])
    expect(r.book.sheets[0].filtered).toBeUndefined()
    // a rule of any value with xlsx's defaults, read as the model's rule with them spelled out
    expect(r.book.sheets[0].validations).toEqual([{ range: 'A1', type: 'any', allowBlank: false, showPrompt: false, showError: false }])
  })

  it('the locale from its custom property; an unknown one reported; the default font other than Calibri 11 reported', async () => {
    const custom = (tag: string) => ({
      '_rels/.rels': rels(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="${REL}/custom-properties" Target="docProps/custom.xml"/>`),
      'docProps/custom.xml': `<Properties xmlns:vt="v"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="xlcalc.locale"><vt:lpwstr>${tag}</vt:lpwstr></property></Properties>`,
    })
    expect((await read({ sheets: [{ name: 'S', body: one(1) }], parts: custom('ru-RU') })).book.locale).toBe('ru-RU')
    const unknown = await read({ sheets: [{ name: 'S', body: one(1) }], parts: custom('de-DE'), styles: '<fonts count="1"><font><sz val="10"/><name val="Arial"/></font></fonts>' })
    expect(unknown.book.locale).toBeUndefined()
    expect(unknown.issues.map((i) => i.message)).toEqual(["the default font Arial 10 is read as Calibri 11, the model's default", 'the locale "de-DE" is not one of the package\'s locales and is left out'])
  })

  it('a sheet without sheetData is an empty sheet', async () => {
    const r = await sheet('')
    expect(r.book.sheets).toEqual([{ name: 'S', cells: {} }])
  })
})
