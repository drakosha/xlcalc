/**
 * The reader on formulas: the file's spelling turned into the engine's (prefixes, spill references, the @ operator),
 * shared formulas spread to their cells as copyRange pastes them, array formulas of both kinds, formulas of the old
 * kind with `@` where Excel 365 puts it, and what the engine cannot take (user functions, other workbooks)
 * reported on its cells.
 */
import { describe, expect, it } from 'vitest'
import { fromFile } from '../src/xlsx/formulas/from-file.js'
import { codes, read, sheet, storedValues } from './xlsx-fixtures.js'

const META_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata'
const METADATA =
  '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>'

/** A sheet with the dynamic array metadata, as Excel 365 writes it. */
const dynamic = (rows: string) =>
  read({
    sheets: [{ name: 'S', body: `<sheetData>${rows}</sheetData>` }],
    parts: {
      'xl/metadata.xml': METADATA,
      'xl/_rels/workbook.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId9" Type="${META_REL}" Target="metadata.xml"/></Relationships>`,
    },
  })

describe("xlsx reader: formulas in the engine's spelling", () => {
  it('prefixes of newer functions and of LET names go; text in quotes stays', () => {
    expect(fromFile('_xlfn.XLOOKUP(1,A:A,B:B)+_xlfn._xlws.FILTER(A1:A3,A1:A3>0)').formula).toBe('XLOOKUP(1,A:A,B:B)+FILTER(A1:A3,A1:A3>0)')
    expect(fromFile('_xlfn.LET(_xlpm.x,1,_xlpm.x+1)').formula).toBe('LET(x,1,x+1)')
    expect(fromFile('"_xlfn.X"&\'_xlfn.Sheet\'!A1&my_xlfn.name').formula).toBe('"_xlfn.X"&\'_xlfn.Sheet\'!A1&my_xlfn.name')
    expect(fromFile('_XLFN.CONCAT("a")').formula).toBe('CONCAT("a")')
  })

  it('ANCHORARRAY is the spill reference; one that does not hold a cell stays a call', () => {
    expect(fromFile('SUM(_xlfn.ANCHORARRAY(B2))*_xlfn.ANCHORARRAY(\'My sheet\'!$C$3)').formula).toBe("SUM(B2#)*'My sheet'!$C$3#")
    expect(fromFile('_xlfn.ANCHORARRAY(#REF!)').formula).toBe('#REF!#')
    expect(fromFile('_xlfn.ANCHORARRAY(A1:B2)').formula).toBe('ANCHORARRAY(A1:B2)')
  })

  it('SINGLE is the @ operator, over a parenthesis where @ would not take the whole argument', () => {
    expect(fromFile('_xlfn.SINGLE(A1:A3)*2+_xlfn.SINGLE(_xlfn.SEQUENCE(3))').formula).toBe('@A1:A3*2+@SEQUENCE(3)')
    expect(fromFile('_xlfn.SINGLE(A1:A3*2)').formula).toBe('@(A1:A3*2)')
    expect(fromFile("_xlfn.SINGLE('My sheet'!B1:B3)").formula).toBe("@'My sheet'!B1:B3")
  })

  it('what the engine has no spelling for is found; references across sheets are read', () => {
    expect(fromFile('_xludf.MYFN(1)')).toEqual({ formula: 'MYFN(1)', found: ['userFunction'] })
    expect(fromFile('SUM(Table1[Amount])+Table1[[#This Row],[Qty]]')).toEqual({ formula: 'SUM(Table1[Amount])+Table1[[#This Row],[Qty]]', found: [] })
    expect(fromFile('SUM(T[@at])+T[[#Headers],[@b]]').formula).toBe("SUM(T['@at])+T[[#Headers],['@b]]")
    expect(fromFile('[@Qty]*2').found).toEqual(['table'])
    expect(fromFile('[1]Sheet1!A1+\'[2]My sheet\'!B2').found).toEqual(['externalLink'])
    expect(fromFile('[0]!rate*2')).toEqual({ formula: 'rate*2', found: [] })
    expect(fromFile("SUM(Sheet1:Sheet3!A1)+SUM('Jan 1:Mar 2'!A1)")).toEqual({ formula: "SUM(Sheet1:Sheet3!A1)+SUM('Jan 1:Mar 2'!A1)", found: [] })
    // LET takes a cell's word as a name but one that reads as R1C1's row or column too, and a boolean (entry/let-cells-held)
    expect(fromFile('LET(_xlpm.r2,1,_xlpm.r2)').found).toEqual(['letName'])
    expect(fromFile('LET(_xlpm.TRUE,1,_xlpm.TRUE)').found).toEqual(['letName'])
    expect(fromFile('LET(_xlpm.a1,1,_xlpm.a1)').found).toEqual([])
    expect(fromFile('LET(_xlpm.R1?x,1,_xlpm.R1?x)')).toEqual({ formula: 'LET(R1?x,1,R1?x)', found: [] })
    expect(fromFile('SUM(Sheet1!A1:B2)+"Sheet1:Sheet3!"').found).toEqual([])
  })
})

describe('xlsx reader: formulas of the file', () => {
  it('a shared formula is spread to its cells as a copy and paste spreads it, a master below and right of copies too', async () => {
    const rows =
      '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2+$A$1+Other!A1</f><v>2</v></c></row>' +
      '<row r="2"><c r="B2" s="0"><f t="shared" si="0"/><v>0</v></c></row><row r="3"><c r="B3"><f t="shared" si="0"/><v>0</v></c></row>' +
      // a master that a copy precedes, and a copy that lands past the sheet's top
      '<row r="4"><c r="D4"><f t="shared" si="1"/></c></row><row r="5"><c r="E5"><f t="shared" ref="D4:E5" si="1">A1+B2</f></c></row>'
    const r = await read({ sheets: [{ name: 'S', body: `<sheetData>${rows}</sheetData>` }, { name: 'Other', body: '<sheetData/>' }] })
    expect(r.book.sheets[0].cells).toMatchObject({
      B1: { formula: 'A1*2+$A$1+Other!A1' },
      B2: { formula: 'A2*2+$A$1+Other!A2' },
      B3: { formula: 'A3*2+$A$1+Other!A3' },
      E5: { formula: 'A1+B2' },
      D4: { formula: '#REF!+A1' },
    })
    expect(r.issues).toEqual([])
  })

  it('a dynamic array formula: the cells it spilled into hold stored values, keeping their own look', async () => {
    const r = await dynamic('<row r="1"><c r="A1" cm="1"><f t="array" ref="A1:A3">_xlfn.SEQUENCE(3)</f><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row><row r="3"><c r="A3" s="0"><v>3</v></c></row>')
    expect(r.book.sheets[0].cells).toEqual({ A1: { formula: 'SEQUENCE(3)' } })
    expect(storedValues(r)).toEqual({ A1: 1, A2: 2, A3: 3 })
    expect(r.issues).toEqual([])
  })

  it('an error an array spilled is a stored value, not an error typed into the cell', async () => {
    const r = await dynamic('<row r="1"><c r="A1" cm="1"><f t="array" ref="A1:B1">{1,#N/A}</f><v>1</v></c><c r="B1" t="e"><v>#N/A</v></c><c r="C1" t="e"><v>#N/A</v></c></row>')
    expect(r.book.sheets[0].cells).toEqual({ A1: { formula: '{1,#N/A}' }, C1: { formula: '#N/A' } })
    expect(storedValues(r)).toEqual({ A1: 1, B1: { error: 'NA' }, C1: { error: 'NA' } })
    expect(codes(r.issues)).toEqual(['errorConstant C1'])
  })

  it('the cells a volatile array spilled into carry an empty formula that only flags them (<f ca="1"/>): stored values', async () => {
    const r = await dynamic('<row r="1"><c r="A1" cm="1"><f t="array" aca="1" ref="A1:A2" ca="1">_xlfn.RANDARRAY(2)</f><v>0.5</v></c></row><row r="2"><c r="A2"><f ca="1"/><v>0.25</v></c></row>')
    expect(r.book.sheets[0].cells).toEqual({ A1: { formula: 'RANDARRAY(2)' } })
    expect(storedValues(r)).toEqual({ A1: 0.5, A2: 0.25 })
    expect(r.issues).toEqual([])
  })

  it('an array formula of the old kind is read over its area, without @', async () => {
    const r = await sheet('<row r="1"><c r="B1"><f t="array" ref="B1:B2">A1:A2*2</f><v>2</v></c><c r="C1"><f t="array" ref="C1">SUM(A1:A2*2)</f><v>6</v></c></row><row r="2"><c r="B2"><v>4</v></c></row>')
    expect(r.book.sheets[0].cells).toEqual({ B1: { formula: 'A1:A2*2', array: 'B1:B2' }, C1: { formula: 'SUM(A1:A2*2)', array: 'C1' } })
    expect(storedValues(r)).toEqual({ B1: 2, B2: 4, C1: 6 })
    expect(r.issues).toEqual([])
  })

  it('a formula of the old kind gets @ where Excel 365 shows it (fromLegacy); one that needs none keeps its text', async () => {
    const formulas = ['A1:A3*2', 'ABS(A1:A3)', 'SUM(A1:A3)', 'SUM(IF(A1:A3>0,1,0))', 'SUM(IF(TRUE,A1:A3))', 'INDEX(A1:B5,0,1)', 'COUNTIF(A:A,">0")', 'A:A', 'ROW()', 'IF(A1>0,1,2)', 'VLOOKUP(1,A1:B3,2,FALSE)', 'span*1', 'one*1', '{1,2}+0', 'sum( A1:A3 )']
    const rows = formulas.map((f, i) => `<row r="${i + 1}"><c r="C${i + 1}"><f>${f.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</f></c></row>`).join('')
    const r = await read({
      sheets: [{ name: 'S', body: `<sheetData>${rows}</sheetData>` }],
      workbook: { after: "<definedNames><definedName name=\"span\">S!$A$1:$A$3</definedName><definedName name=\"one\">S!$A$1</definedName></definedNames>" },
    })
    const shown = ['@A1:A3*2', 'ABS(@A1:A3)', 'SUM(A1:A3)', 'SUM(IF(@A1:A3>0,1,0))', 'SUM(IF(TRUE,A1:A3))', '@INDEX(A1:B5,0,1)', 'COUNTIF(A:A,">0")', '@A:A', 'ROW()', 'IF(A1>0,1,2)', 'VLOOKUP(1,A1:B3,2,FALSE)', '@span*1', 'one*1', '@{1,2}+0', 'sum( A1:A3 )']
    expect(formulas.map((_, i) => r.book.sheets[0].cells[`C${i + 1}`].formula)).toEqual(shown)
    expect(r.issues).toEqual([])
  })

  it('shared formulas of the old kind: each copy with @ as its master', async () => {
    const rows = '<row r="1"><c r="B1"><f t="shared" ref="B1:B2" si="0">A1:A3*2</f></c></row><row r="2"><c r="B2"><f t="shared" si="0"/></c></row>'
    const r = await sheet(rows)
    expect(r.book.sheets[0].cells).toEqual({ B1: { formula: '@A1:A3*2' }, B2: { formula: '@A2:A4*2' } })
  })

  it('formulas with dynamic array metadata are not of the old kind', async () => {
    const r = await dynamic('<row r="1"><c r="C1" cm="1"><f t="array" ref="C1:C3">A1:A3*2</f></c></row>')
    expect(r.book.sheets[0].cells).toEqual({ C1: { formula: 'A1:A3*2' } })
    expect(r.issues).toEqual([])
  })

  it('what the engine cannot take is reported on its cells, the formula kept as written', async () => {
    const r = await sheet(
      '<row r="1"><c r="A1"><f>_xlfn.SINGLE(B1:B3)</f></c><c r="B1"><f>_xludf.MYFN(1)</f></c><c r="C1"><f>SUM(T[x])</f></c><c r="D1"><f>[1]S!A1</f></c><c r="E1"><f>SUM(S:T!A1)</f></c><c r="F1"><f>A1 B1</f></c><c r="G1"><f>SUM(</f></c></row>',
    )
    expect(r.book.sheets[0].cells).toMatchObject({ A1: { formula: '@B1:B3' }, B1: { formula: 'MYFN(1)' }, C1: { formula: 'SUM(T[x])' }, D1: { formula: '[1]S!A1' }, E1: { formula: 'SUM(S:T!A1)' } })
    expect(r.book.sheets[0].cells.F1).toEqual({ formula: 'A1 B1' })
    expect(codes(r.issues)).toEqual(['userFunction B1', 'externalLink D1', 'formula G1'])
  })
})
