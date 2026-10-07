/**
 * The reader on formulas: the file's spelling turned into the engine's (prefixes, spill references), shared
 * formulas spread to their cells as copyRange pastes them, array formulas of both kinds, implicit intersection
 * in formulas of the old kind, and what the engine cannot take (the @ operator, user functions, tables, other
 * workbooks, references across sheets) reported on its cells.
 */
import { describe, expect, it } from 'vitest'
import { fromFile } from '../src/xlsx/formulas/from-file.js'
import { codes, read, sheet } from './xlsx-fixtures.js'

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

  it('what the engine has no spelling for is found', () => {
    expect(fromFile('_xlfn.SINGLE(A1:A3)').found).toEqual(['singleOperator'])
    expect(fromFile('_xludf.MYFN(1)')).toEqual({ formula: 'MYFN(1)', found: ['userFunction'] })
    expect(fromFile('SUM(Table1[Amount])+Table1[[#This Row],[Qty]]').found).toEqual(['table'])
    expect(fromFile('[1]Sheet1!A1+\'[2]My sheet\'!B2').found).toEqual(['externalLink'])
    expect(fromFile('SUM(Sheet1:Sheet3!A1)').found).toEqual(['reference3d'])
    expect(fromFile("SUM('Jan 1:Mar 2'!A1)").found).toEqual(['reference3d'])
    expect(fromFile('LET(_xlpm.a1,1,_xlpm.a1)').found).toEqual(['letName'])
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
    expect(r.cached[0].cells).toEqual({ A1: 1, A2: 2, A3: 3 })
    expect(r.issues).toEqual([])
  })

  it('an array formula of the old kind is read as a dynamic one, and reported', async () => {
    const r = await sheet('<row r="1"><c r="B1"><f t="array" ref="B1:B2">A1:A2*2</f><v>2</v></c></row><row r="2"><c r="B2"><v>4</v></c></row>')
    expect(r.book.sheets[0].cells).toEqual({ B1: { formula: 'A1:A2*2' } })
    expect(r.cached[0].cells).toEqual({ B1: 2, B2: 4 })
    expect(codes(r.issues)).toEqual(['legacyArray B1:B2', 'implicitIntersection B1'])
  })

  it('implicit intersection: a range where one value goes, in a formula of the old kind; not where the function takes it whole', async () => {
    const formulas = ['A1:A3*2', 'ABS(A1:A3)', 'SUM(A1:A3)', 'SUM(IF(A1:A3>0,1,0))', 'SUM(IF(TRUE,A1:A3))', 'INDEX(A1:B5,0,1)', 'COUNTIF(A:A,">0")', 'A:A', 'ROW()', 'IF(A1>0,1,2)', 'VLOOKUP(1,A1:B3,2,FALSE)', 'span*1', 'one*1', '{1,2}+0']
    const rows = formulas.map((f, i) => `<row r="${i + 1}"><c r="C${i + 1}"><f>${f.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</f></c></row>`).join('')
    const r = await read({
      sheets: [{ name: 'S', body: `<sheetData>${rows}</sheetData>` }],
      workbook: { after: "<definedNames><definedName name=\"span\">S!$A$1:$A$3</definedName><definedName name=\"one\">S!$A$1</definedName></definedNames>" },
    })
    expect(codes(r.issues)).toEqual(['implicitIntersection C1 C2 C4 C6 C8 C12 C14'])
  })

  it('formulas with dynamic array metadata are not of the old kind', async () => {
    const r = await dynamic('<row r="1"><c r="C1" cm="1"><f t="array" ref="C1:C3">A1:A3*2</f></c></row>')
    expect(r.issues).toEqual([])
  })

  it('what the engine cannot take is reported on its cells, the formula kept as written', async () => {
    const r = await sheet(
      '<row r="1"><c r="A1"><f>_xlfn.SINGLE(B1:B3)</f></c><c r="B1"><f>_xludf.MYFN(1)</f></c><c r="C1"><f>SUM(T[x])</f></c><c r="D1"><f>[1]S!A1</f></c><c r="E1"><f>SUM(S:T!A1)</f></c><c r="F1"><f>A1 B1</f></c><c r="G1"><f>SUM(</f></c></row>',
    )
    expect(r.book.sheets[0].cells).toMatchObject({ A1: { formula: 'SINGLE(B1:B3)' }, B1: { formula: 'MYFN(1)' }, C1: { formula: 'SUM(T[x])' }, D1: { formula: '[1]S!A1' } })
    expect(codes(r.issues)).toEqual(['singleOperator A1', 'userFunction B1', 'table C1', 'externalLink D1', 'reference3d E1', 'formula F1 G1'])
  })
})
