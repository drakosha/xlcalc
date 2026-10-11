/**
 * Styles both ways: theme colours with their tint, the legacy palette and a book's own, built-in and custom
 * number formats, equal styles shared, fonts relative to the default one, and what a style cannot hold reported
 * on the cells that use it. The writer's styles read back as written.
 */
import { describe, expect, it } from 'vitest'
import type { BookInput } from '../src/contract.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { palette, tinted } from '../src/xlsx/styles/color.js'
import { readFill, underlineOf } from '../src/xlsx/styles/read.js'
import { element } from '../src/xlsx/xml/scan.js'
import type { StyleLoss } from '../src/locale/messages.js'
import { codes, read } from './xlsx-fixtures.js'

const THEME =
  '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>'

const withTheme = (styles: string, cells: string) =>
  read({
    sheets: [{ name: 'S', body: `<sheetData><row r="1">${cells}</row></sheetData>` }],
    styles,
    parts: {
      'xl/theme/theme1.xml': THEME,
      'xl/_rels/workbook.xml.rels':
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rIdY" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>',
    },
  })

const FONTS = '<fonts count="2"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><scheme val="minor"/></font><font><b/><i val="0"/><sz val="11"/><color theme="4" tint="0.39997558519241921"/><name val="Calibri"/></font></fonts>'

describe('xlsx styles: reading', () => {
  it("theme colours: the theme's index swaps its first two pairs; a tint lightens or darkens by luminance", async () => {
    const fills =
      '<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor theme="0" tint="-0.249977111117893"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor indexed="10"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="00FFC7CE"/></patternFill></fill></fills>'
    const xfs = '<cellXfs count="5"><xf fontId="0"/><xf fontId="1"/><xf fillId="2"/><xf fillId="3"/><xf fillId="4"/></cellXfs>'
    const r = await withTheme(FONTS + fills + xfs, '<c r="A1" s="1"><v>1</v></c><c r="B1" s="2"><v>1</v></c><c r="C1" s="3"><v>1</v></c><c r="D1" s="4"><v>1</v></c><c r="E1" s="0"><v>1</v></c>')
    expect(r.book.styles).toEqual([{ font: { bold: true, color: '#8EA9DB' } }, { fill: { color: '#BFBFBF' } }, { fill: { color: '#FF0000' } }, { fill: { color: '#FFC7CE' } }])
    expect(r.book.sheets[0].cells).toEqual({ A1: { value: 1, style: 0 }, B1: { value: 1, style: 1 }, C1: { value: 1, style: 2 }, D1: { value: 1, style: 3 }, E1: { value: 1 } })
    expect(r.issues).toEqual([])
  })

  it("tint on luminance in Windows' HLS steps, as Excel shows it (the colours Excel showed: test/excel/xlsx/colors.json)", () => {
    expect(tinted('4472C4', 0.39997558519241921)).toBe('8EA9DB')
    expect(tinted('4472C4', -0.249977111117893)).toBe('305496')
    expect(tinted('FFFFFF', -0.0499893185216834)).toBe('F2F2F2')
    expect(tinted('A5A5A5', -0.499984740745262)).toBe('525252')
    expect(tinted('ABCDEF', 0)).toBe('ABCDEF')
  })

  it("a book's own palette replaces the legacy one; auto colours are no colour", async () => {
    const r = await read({
      sheets: [{ name: 'S', body: '<sheetData><row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="2"><v>1</v></c></row></sheetData>' }],
      styles: '<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><sz val="11"/><color indexed="2"/><name val="Calibri"/></font><font><sz val="11"/><color auto="1"/><name val="Calibri"/></font></fonts><cellXfs count="3"><xf/><xf fontId="1"/><xf fontId="2"/></cellXfs><colors><indexedColors><rgbColor rgb="FF000000"/><rgbColor rgb="FFFFFFFF"/><rgbColor rgb="FF123456"/></indexedColors></colors>',
    })
    expect(r.book.styles).toEqual([{ font: { color: '#123456' } }])
    expect(r.book.sheets[0].cells).toEqual({ A1: { value: 1, style: 0 }, B1: { value: 1 } })
  })

  it('fonts relative to the default, borders, alignment; equal styles shared; built-in and custom formats', async () => {
    const r = await read({
      sheets: [{ name: 'S', body: '<sheetData><row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="2"><v>1</v></c><c r="C1" s="3"><v>1</v></c><c r="D1" s="4"><v>1</v></c></row></sheetData>' }],
      styles:
        '<numFmts count="1"><numFmt numFmtId="170" formatCode="#,##0.00&quot; ₽&quot;"/></numFmts>' +
        '<fonts count="2"><font><sz val="10"/><name val="Arial"/></font><font><u val="double"/><strike/><sz val="12"/><name val="Arial"/></font></fonts>' +
        '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FF00B050"/></left><right style="medium"/><top/><bottom style="double"/><diagonal/></border></borders>' +
        '<cellXfs count="5"><xf/><xf fontId="1" numFmtId="170"/><xf borderId="1" numFmtId="3"><alignment horizontal="center" vertical="top" wrapText="1" indent="2"/></xf><xf fontId="1" numFmtId="49"/><xf borderId="1" numFmtId="14"><alignment horizontal="center" vertical="top" wrapText="1" indent="2"/></xf></cellXfs>',
    })
    expect(r.book.styles).toEqual([
      { font: { underline: 'double', strike: true, size: 12 } },
      { border: { left: { style: 'thin', color: '#00B050' }, right: { style: 'medium' }, bottom: { style: 'double' } }, alignment: { horizontal: 'center', vertical: 'top', wrap: true, indent: 2 } },
    ])
    expect(r.book.sheets[0].cells).toEqual({ A1: { value: 1, style: 0, numberFormat: '#,##0.00" ₽"' }, B1: { value: 1, style: 1, numberFormat: 3 }, C1: { value: 1, style: 0, numberFormat: 49 }, D1: { value: 1, style: 1, numberFormat: 14 } })
    expect(codes(r.issues)).toEqual(['defaultFont'])
  })

  it('what a style cannot hold is reported on its cells; an East Asian date format is read as its en-US code, an unknown built-in format is left out', async () => {
    const r = await read({
      sheets: [{ name: 'S', body: '<sheetData><row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="2"><v>1</v></c><c r="C1" s="1"><v>1</v></c><c r="D1" s="3"><v>1</v></c><c r="E1" s="4"><v>1</v></c></row></sheetData>' }],
      styles:
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><vertAlign val="superscript"/><sz val="11"/><name val="Calibri"/></font></fonts>' +
        '<fills count="4"><fill><patternFill/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="darkGrid"><fgColor rgb="FFFF0000"/></patternFill></fill><fill><gradientFill><stop position="0"><color rgb="FF000000"/></stop></gradientFill></fill></fills>' +
        '<borders count="2"><border/><border diagonalUp="1"><diagonal style="thin"/></border></borders>' +
        '<cellXfs count="5"><xf/><xf fontId="1" fillId="2"><alignment textRotation="45" shrinkToFit="1"/><protection locked="0"/></xf><xf fillId="3" borderId="1"/><xf numFmtId="55"/><xf numFmtId="81"/></cellXfs>',
    })
    expect(r.book.styles).toBeUndefined()
    expect(r.book.sheets[0].cells.D1).toEqual({ value: 1, numberFormat: 'm/d/yy' })
    expect(r.book.sheets[0].cells.E1).toEqual({ value: 1 })
    expect(codes(r.issues)).toEqual([
      'style A1 C1',
      'style A1 C1',
      'style A1 C1',
      'style A1 C1',
      'style A1 C1',
      'style B1',
      'style B1',
      'numberFormat D1',
      'numberFormat E1',
    ])
    expect(r.issues.filter((i) => i.code === 'style').map((i) => i.message)).toEqual([
      "a style's superscript is not in the model",
      "a style's pattern fill is not in the model",
      "a style's text rotation is not in the model",
      "a style's shrink to fit is not in the model",
      "a style's protection is not in the model",
      "a style's gradient fill is not in the model",
      "a style's diagonal border is not in the model",
    ])
  })

  it("a cell's font and a dxf's: an accounting underline is read as its plain kind, lost; a dxf's none takes the underline away, lost too", () => {
    const under = (xml: string, dxf = false) => {
      const lost: StyleLoss[] = []
      return [underlineOf(xml, lost, dxf), lost]
    }
    expect(under('<u/>')).toEqual(['single', []])
    expect(under('<u val="double"/>', true)).toEqual(['double', []])
    expect(under('<u val="singleAccounting"/>')).toEqual(['single', ['accountingUnderline']])
    expect(under('<u val="doubleAccounting"/>', true)).toEqual(['double', ['accountingUnderline']])
    expect(under('<u val="none"/>')).toEqual([undefined, []])
    expect(under('<u val="none"/>', true)).toEqual([undefined, ['underlineRemoved']])
    expect(under('<b/>', true)).toEqual([undefined, []])
  })

  it("a fill as Excel shows it (cond/fills-held, fills-held-2): a cell's solid fgColor, white without one, none without a type; a dxf's bgColor alone, solid without a type, none taking the fill away", () => {
    const p = palette(undefined, undefined)
    const fill = (body: string, dxf = false) => readFill(element(`<fill>${body}</fill>`, 'fill')!, p, dxf)
    const both = '<fgColor rgb="FF00FF00"/><bgColor rgb="FFFF0000"/>'
    const gradient = '<gradientFill degree="90"><stop position="0"><color rgb="FFFF0000"/></stop></gradientFill>'
    expect([fill(`<patternFill patternType="solid">${both}</patternFill>`), fill(`<patternFill patternType="solid">${both}</patternFill>`, true)]).toEqual([{ fill: { color: '#00FF00' }, lost: [] }, { fill: { color: '#FF0000' }, lost: [] }])
    expect(fill('<patternFill patternType="solid"><bgColor rgb="FFFF0000"/></patternFill>')).toEqual({ fill: { color: '#FFFFFF' }, lost: [] })
    expect(fill('<patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill>', true)).toEqual({ lost: ['fillWithoutColor'] })
    // a colour that names none is white, in either; a dxf's empty pattern fill changes nothing (cond/fills-held-2)
    expect([fill('<patternFill patternType="solid"><fgColor auto="1"/><bgColor rgb="FFFF0000"/></patternFill>'), fill('<patternFill><bgColor indexed="64"/></patternFill>', true)]).toEqual([{ fill: { color: '#FFFFFF' }, lost: [] }, { fill: { color: '#FFFFFF' }, lost: [] }])
    expect(fill('<patternFill/>', true)).toEqual({ lost: [] })
    expect([fill('<patternFill><fgColor rgb="FFFF0000"/></patternFill>'), fill('<patternFill><bgColor rgb="FFFF0000"/></patternFill>', true)]).toEqual([{ lost: [] }, { fill: { color: '#FF0000' }, lost: [] }])
    expect([fill('<patternFill patternType="none"><bgColor rgb="FFFF0000"/></patternFill>'), fill('<patternFill patternType="none"><bgColor rgb="FFFF0000"/></patternFill>', true)]).toEqual([{ lost: [] }, { lost: ['fillRemoved'] }])
    expect([fill(gradient), fill(gradient, true)]).toEqual([{ lost: ['gradientFill'] }, { lost: ['gradientFill'] }])
    expect(fill(`${gradient}<patternFill patternType="solid">${both}</patternFill>`)).toEqual({ fill: { color: '#00FF00' }, lost: ['gradientBesidePattern'] })
    expect(fill('<patternFill patternType="darkGray"><fgColor rgb="FF00FF00"/></patternFill>', true)).toEqual({ lost: ['patternFill'] })
  })
})

describe('xlsx styles: writing and reading back', () => {
  const book: BookInput = {
    sheets: [
      {
        name: 'S',
        cells: {
          A1: { value: 1, style: 0, numberFormat: '0.00' }, A2: { value: 2, style: 1 }, A3: { value: 3, numberFormat: 'dd.mm.yyyy' }, A4: { value: 4, numberFormat: 'm/d/yyyy' }, A5: { value: 5, numberFormat: '0%' }, A6: { style: 2 },
          A7: { value: 7, numberFormat: 14 }, A8: { value: -8, numberFormat: 38 }, A9: { value: -9, numberFormat: '#,##0_);(#,##0)' }, A10: { value: 10, numberFormat: 'm/d/yy' },
        },
        columns: { B: { width: 15.5, style: 1 }, 'C:E': { numberFormat: '0.0' } },
        rows: { 2: { height: 25 }, '7:8': { style: 0 } },
      },
    ],
    styles: [
      { font: { name: 'Arial', size: 9, bold: true, italic: true, underline: 'single', strike: true, color: '#1F4E79' } },
      { fill: { color: '#E2EFDA' }, border: { top: { style: 'mediumDashDot', color: '#FF0000' }, left: { style: 'hair' } }, alignment: { horizontal: 'distributed', vertical: 'justify', indent: 3 } },
      { alignment: { wrap: true } },
    ],
  }

  it('styles, formats of every kind, column and row looks come back as written; a code every locale shows as a built-in format comes back as its number', async () => {
    const { data, issues } = await writeXlsx({ book })
    expect(issues).toEqual([])
    const r = await readXlsx(data)
    const cells = { ...book.sheets[0].cells, A1: { ...book.sheets[0].cells.A1, numberFormat: 2 }, A5: { value: 5, numberFormat: 9 } }
    expect(r.book).toEqual({ ...book, sheets: [{ ...book.sheets[0], cells }] })
    expect(r.issues).toEqual([])
  })
})
