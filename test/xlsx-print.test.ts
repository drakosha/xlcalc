import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { layoutPrintPages, printPaperSize, printSheetGeometry, readXlsxWithPrint, writeXlsxWithPrint, type SheetPrint } from '../src/xlsx-print.js'
import { writeXlsx } from '../src/xlsx.js'
import { partsOf, withParts } from './keep-files.js'

describe('xlsx/print', () => {
  it('reads Excel’s print area, margins and manual page break', async () => {
    const bytes = readFileSync(new URL('./excel/keep/open/print-parts.xlsx', import.meta.url))
    const read = await readXlsxWithPrint(bytes)
    expect(read.print.find((settings) => settings.sheet === 'Data')).toMatchObject({
      areas: ['A1:C21'],
      margins: { left: 0.7, top: 0.75 },
      setup: { orientation: 'portrait' },
      rowBreaks: [{ before: 12, last: 3, manual: true }],
    })
  })

  it('uses sheet-specific default height and rows written on a sheet hidden by default', async () => {
    const bytes = await withParts({
      sheets: [{
        name: 'Report', cells: { A2: { value: 2 }, A3: { value: 3 }, A4: { value: 4 } },
        rows: { '3': { hidden: true }, '4': { height: 20 } },
      }],
    }, { edit: {
      'xl/worksheets/sheet1.xml': (xml) => xml.replace('<sheetFormatPr defaultRowHeight="15"', '<sheetFormatPr defaultRowHeight="23" customHeight="1" zeroHeight="1"'),
    } })
    const read = await readXlsxWithPrint(bytes)
    const defaults = read.sheetDefaults[0]
    expect(defaults).toEqual({ sheet: 'Report', defaultRowHeight: 23, visibleRows: [{ first: 2, last: 2 }, { first: 4, last: 4 }] })
    const geometry = printSheetGeometry(read.book.sheets[0], { ...defaults, usedRange: 'A1:A6' })
    expect([1, 2, 3, 4, 5, 6].map(geometry.rows.edge)).toEqual([0, 0, 23, 23, 43, 43])
  })

  it('writes print options in worksheet order and round-trips names scoped by sheet index', async () => {
    const snapshot = {
      book: {
        sheets: [
          { name: 'Before', cells: { A1: { value: 1 } } },
          { name: "Bob's, отчёт", tabColor: '#112233', outline: { summaryBelow: false }, cells: { A1: { value: 2, note: { text: 'Note' } } } },
        ],
        names: [{ name: 'Total', ref: 'Before!$A$1' }],
      },
    }
    const print: SheetPrint[] = [{
      sheet: "Bob's, отчёт",
      areas: ['A1:B30', 'D1:E30'],
      repeatRows: { first: 1, last: 2 },
      repeatColumns: { first: 1, last: 1 },
      setup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 },
      margins: { left: 0.5, right: 0.5, top: 0.6, bottom: 0.6, header: 0.2, footer: 0.2 },
      options: { horizontalCentered: true, gridLines: true, gridLinesSet: true },
      headerFooter: { oddFooter: '&LReport < Q4&R&P / &N', differentFirst: true, firstFooter: '&RFirst' },
      rowBreaks: [{ before: 8, first: 1, last: 2, manual: true }],
      columnBreaks: [{ before: 4, manual: true }],
    }]
    const written = await writeXlsxWithPrint(snapshot, print)
    const parts = await partsOf(written.data)
    const sheet = parts.get('xl/worksheets/sheet2.xml')!
    expect(sheet).toMatch(/<tabColor[^>]*\/><outlinePr[^>]*\/><pageSetUpPr[^>]*\/><\/sheetPr>/)
    expect(sheet).toMatch(/<printOptions[^>]*\/><pageMargins[^>]*\/><pageSetup[^>]*\/><headerFooter[\s\S]*<\/headerFooter><rowBreaks[\s\S]*<\/rowBreaks><colBreaks[\s\S]*<\/colBreaks><legacyDrawing/)
    expect(parts.get('xl/workbook.xml')).toContain('name="_xlnm.Print_Titles" localSheetId="1"')
    expect((await readXlsxWithPrint(written.data)).print).toEqual(print)
    expect((await writeXlsxWithPrint(snapshot, [])).data).toEqual((await writeXlsx(snapshot)).data)
    await expect(writeXlsxWithPrint(snapshot, [{ ...print[0], sheet: 'Missing' }])).rejects.toThrow(/Missing/)
  })

  it('uses Excel paper stock codes and keeps the physical output target a caller choice', () => {
    const a4 = printPaperSize({ paperSize: 9, orientation: 'landscape', paperWidth: '8in', paperHeight: '12in' })!
    expect(a4.width).toBeCloseTo(841.8898)
    expect(a4.height).toBeCloseTo(595.2756)
    const rotatedA4 = printPaperSize({ paperSize: 77, orientation: 'portrait' })!
    expect(rotatedA4.width).toBeCloseTo(a4.height)
    expect(rotatedA4.height).toBeCloseTo(a4.width)
    expect(printPaperSize({ paperSize: 1, orientation: 'landscape' })).toEqual({ width: 792, height: 612 })
    expect(printPaperSize({ paperSize: 256 }, { width: 612, height: 792 })).toEqual({ width: 612, height: 792 })
    expect(printPaperSize({ paperSize: 256 })).toBeUndefined()
  })

  it('round-trips a negative page number as Excel’s unsigned field and accepts automatic scale', async () => {
    const snapshot = { book: { sheets: [{ name: 'Report', cells: { A1: { value: 1 } } }] } }
    const print: SheetPrint = { sheet: 'Report', setup: { firstPageNumber: -2, useFirstPageNumber: true, scale: 0 } }
    const written = await writeXlsxWithPrint(snapshot, [print])
    const parts = await partsOf(written.data)
    expect(parts.get('xl/worksheets/sheet1.xml')).toContain('firstPageNumber="4294967294"')
    expect((await readXlsxWithPrint(written.data)).print).toEqual([print])
    const geometry = printSheetGeometry({}, { usedRange: 'A1' })
    const { pages } = layoutPrintPages(print, geometry, { width: 612, height: 792 })
    expect(pages[0]).toMatchObject({ number: -2, scale: 1 })
  })

  it('calculates page regions in sheet and paper coordinates, repeating headings and skipping hidden lines', () => {
    const rowEdges = [0, 12, 32, 32, 52, 72, 92, 112]
    const plan = layoutPrintPages({
      sheet: 'Report',
      repeatRows: { first: 1, last: 1 },
      repeatColumns: { first: 1, last: 1 },
      rowBreaks: [{ before: 5, manual: true }],
      setup: { pageOrder: 'overThenDown', firstPageNumber: 4, useFirstPageNumber: true },
      margins: { top: 0, bottom: 0, left: 0, right: 0, header: 0, footer: 0 },
    }, {
      usedRange: 'A1:D7',
      rows: { count: 7, edge: (line) => rowEdges[line - 1] },
      columns: { count: 4, edge: (line) => 40 * (line - 1) },
    }, { width: 85, height: 64 })
    expect(plan.pages.map((page) => page.number)).toEqual([4, 5, 6, 7, 8, 9, 10, 11, 12])
    expect(plan.pages[0].regions.map((region) => region.range)).toEqual(['A1:B4'])
    const next = plan.pages[4]
    expect(next.regions.map((region) => region.range)).toEqual(['A1', 'C1', 'A5:A6', 'C5:C6'])
    expect(next.regions[3]).toMatchObject({
      source: { x: 80, y: 52, width: 40, height: 40 },
      target: { x: 40, y: 12, width: 40, height: 40 },
    })
  })
  it('fits a wide area at an integer scale and ignores manual breaks in fit mode', () => {
    const margins = { top: 0, bottom: 0, left: 0, right: 0, header: 0, footer: 0 }
    const plan = layoutPrintPages({
      sheet: 'Report',
      repeatRows: { first: 1, last: 1 },
      rowBreaks: [{ before: 5, manual: true }],
      setup: { fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
      headerFooter: {
        differentFirst: true, firstFooter: '&RFirst', differentOddEven: true, evenFooter: '&REven', scaleWithDoc: false,
      },
      margins,
    }, {
      usedRange: 'A1:C10',
      rows: { count: 10, edge: (line) => 30 * (line - 1) },
      columns: { count: 3, edge: (line) => 60 * (line - 1) },
    }, { width: 100, height: 150 })
    expect(plan.issues).toEqual([])
    expect(plan.pages.map((page) => page.scale)).toEqual([0.55, 0.55])
    expect(plan.pages.map((page) => page.regions.map((region) => region.range))).toEqual([['A1:C9'], ['A1:C1', 'A10:C10']])
    expect(plan.pages.map((page) => page.footer)).toMatchObject([
      { text: '&RFirst', scale: 1 }, { text: '&REven', scale: 1 },
    ])
  })

  it('extends a partial break across the print area it intersects', () => {
    const settings: SheetPrint = {
      sheet: 'Report',
      setup: { pageOrder: 'overThenDown' },
      rowBreaks: [{ before: 3, first: 2, last: 2, manual: true }],
      margins: { left: 0, right: 0, top: 0, bottom: 0, header: 0, footer: 0 },
    }
    const geometry = {
      usedRange: 'A1:D7',
      rows: { count: 7, edge: (line: number) => 20 * (line - 1) },
      columns: { count: 4, edge: (line: number) => 20 * (line - 1) },
    }
    const paper = { width: 45, height: 60 }
    const all = layoutPrintPages(settings, geometry, paper)
    expect(all.issues).toEqual([])
    expect(all.pages.map((page) => page.regions[0].range)).toEqual([
      'A1:B2', 'C1:D2', 'A3:B5', 'C3:D5', 'A6:B7', 'C6:D7',
    ])
    const areas = layoutPrintPages({ ...settings, areas: ['A1:B7', 'C1:D7'] }, geometry, paper)
    expect(areas.pages.map((page) => page.regions[0].range)).toEqual([
      'A1:B2', 'A3:B5', 'A6:B7', 'C1:D3', 'C4:D6', 'C7:D7',
    ])
  })

  it('starts repeating print titles after the body reaches their rows', () => {
    const geometry = printSheetGeometry({}, { usedRange: 'A1:A60', defaultRowHeight: 15 })
    const paper = { width: 100, height: 150 }
    const settings: SheetPrint = {
      sheet: 'Report', repeatRows: { first: 20, last: 22 },
      margins: { left: 0, right: 0, top: 0, bottom: 0, header: 0, footer: 0 },
    }
    const above = layoutPrintPages({ ...settings, areas: ['A1:A18'] }, geometry, paper)
    expect(above.pages.map((page) => page.regions.map((region) => region.range))).toEqual([
      ['A1:A10'], ['A11:A18'],
    ])
    const across = layoutPrintPages({ ...settings, areas: ['A1:A39'] }, geometry, paper)
    expect(across.pages.map((page) => page.regions.map((region) => region.range))).toEqual([
      ['A1:A10'], ['A11:A20'], ['A20:A29'], ['A20:A22', 'A30:A36'], ['A20:A22', 'A37:A39'],
    ])
  })

  it('selects even footers by printed position when numbering starts at 4', () => {
    const plan = layoutPrintPages({
      sheet: 'Report',
      setup: { firstPageNumber: 4, useFirstPageNumber: true },
      margins: { left: 0, right: 0, top: 0, bottom: 0, header: 0, footer: 0 },
      headerFooter: { differentFirst: true, differentOddEven: true, firstFooter: 'FIRST', oddFooter: 'ODD', evenFooter: 'EVEN' },
    }, {
      usedRange: 'A1:A4',
      rows: { count: 4, edge: (line) => (line - 1) * 20 },
      columns: { count: 1, edge: (line) => (line - 1) * 20 },
    }, { width: 30, height: 30 })
    expect(plan.pages.map((page) => [page.number, page.footer?.text])).toEqual([
      [4, 'FIRST'], [5, 'EVEN'], [6, 'ODD'], [7, 'EVEN'],
    ])
  })

  it('records the pagination gap from the Windows Excel PDF control', () => {
    const settings: SheetPrint = {
      sheet: 'Ledger',
      areas: ['A1:A125'],
      repeatRows: { first: 1, last: 2 },
      margins: { left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 },
      setup: { fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    }
    const layout = (pitch: number) => layoutPrintPages(settings, {
      rows: { count: 125, edge: (line) => pitch * (line - 1) },
      columns: { count: 1, edge: (line) => 48 * (line - 1) },
    }, { width: 841.92, height: 595.32 }).pages.map((page) => page.regions.at(-1)?.range)
    // Excel ends the first three pages at 35, 68 and 101. Its reported 15 pt worksheet rows
    // and its measured 14.523 pt PDF row pitch put those boundaries on opposite sides.
    expect(layout(15)).toEqual(['A1:A34', 'A35:A66', 'A67:A98', 'A99:A125'])
    expect(layout(14.523)).toEqual(['A1:A36', 'A37:A70', 'A71:A104', 'A105:A125'])
  })

  it('keeps source coordinates while laying out both axes with measured printer geometry', () => {
    const settings: SheetPrint = {
      sheet: 'Ledger',
      areas: ['A1:D125'],
      repeatRows: { first: 1, last: 2 },
      margins: { left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 },
      setup: { fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    }
    const geometry = {
      rows: { count: 125, edge: (line: number) => 15 * (line - 1) },
      columns: { count: 4, edge: (line: number) => 262.5 * (line - 1) },
    }
    const target = {
      width: 841.92, height: 595.32,
      paginationReserve: { height: 0.72 },
      deviceGeometry: {
        rowEdge: (line: number, scale: number) => Math.round(121 * scale) * (line - 1) * 0.12,
        columnEdge: (line: number, scale: number) => Math.round(2350 * scale) * (line - 1) * 0.12,
      },
    }
    const plan = layoutPrintPages(settings, geometry, target)
    expect(plan.issues).toEqual([])
    expect(plan.pages.map((page) => page.scale)).toEqual([0.68, 0.68, 0.68])
    expect(plan.pages.map((page) => page.regions.map((region) => region.range))).toEqual([
      ['A1:D53'], ['A1:D2', 'A54:D104'], ['A1:D2', 'A105:D125'],
    ])
    expect(plan.pages[0].regions[0].source).toEqual({ x: 0, y: 0, width: 1050, height: 795 })
    expect(plan.pages[0].regions[0].target).toMatchObject({ x: 36, y: 36 })
    expect(plan.pages[0].regions[0].target.width).toBeCloseTo(767.04)
    expect(plan.pages[0].regions[0].target.height).toBeCloseTo(521.52)

    const fixed = layoutPrintPages({ ...settings, areas: ['A1:A125'], setup: { scale: 100 } }, geometry, target)
    expect(fixed.pages.map((page) => page.regions.at(-1)?.range)).toEqual(['A1:A35', 'A36:A68', 'A69:A101', 'A102:A125'])
  })

  it('converts workbook line spans into sheet coordinates and bounds a whole-column area to the print extent', () => {
    const geometry = printSheetGeometry({
      defaultColumnWidth: 10,
      columns: { A: { width: 20 }, B: { hidden: true } },
      rows: { '2:3': { height: 25 }, '5': { hidden: true } },
    }, { usedRange: 'A1:C6', defaultRowHeight: 15 })
    expect([1, 2, 3, 4].map(geometry.columns.edge)).toEqual([0, 120, 120, 180])
    expect([1, 2, 3, 4, 5, 6, 7].map(geometry.rows.edge)).toEqual([0, 15, 40, 65, 80, 80, 95])
    const plan = layoutPrintPages({ sheet: 'Report', areas: ['B:C'] }, geometry, { width: 612, height: 792 })
    expect(plan.pages[0].regions).toMatchObject([{ range: 'B1:C6', source: { x: 120, width: 60, height: 95 } }])
  })
})
