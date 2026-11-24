import { describe, expect, it } from 'vitest'
import { formatPrintHeaderFooter, layoutPrintWorkbook, type PrintWorkbookSheet } from '../src/xlsx-print.js'

function sheet(name: string, count: number, firstPageNumber?: number): PrintWorkbookSheet {
  const body = { width: 50, height: 60 }
  return {
    settings: {
      sheet: name,
      setup: firstPageNumber === undefined ? undefined : { firstPageNumber, useFirstPageNumber: true },
      margins: { left: 0, right: 0, top: 0, bottom: 0, header: 0, footer: 0 },
      headerFooter: {
        differentFirst: true, differentOddEven: true,
        firstFooter: '&CFIRST &A &P of &N', evenFooter: '&CEVEN &A &P of &N', oddFooter: '&CODD &A &P of &N',
      },
    },
    geometry: {
      usedRange: 'A1:A' + count,
      rows: { count, edge: (line) => (line - 1) * 20 },
      columns: { count: 1, edge: (line) => (line - 1) * 20 },
    },
    target: body,
  }
}

describe('workbook printing', () => {
  it('chooses headers by physical job position and shares the total across visible sheets', () => {
    const plan = layoutPrintWorkbook([sheet('Report', 2), { ...sheet('Hidden', 7), hidden: 'very' }, sheet('Other', 7)])
    expect(plan.pages.map((page) => [page.sheet, page.position, page.number])).toEqual([
      ['Report', 1, 1], ['Other', 2, 2], ['Other', 3, 3], ['Other', 4, 4],
    ])
    expect(plan.pages.map((page) => formatPrintHeaderFooter(page.footer!.text, {
      sheetName: page.sheet, pageNumber: page.number, pageCount: plan.pages.length,
    }).center[0])).toEqual([
      { text: 'FIRST Report 1 of 4', style: {} },
      { text: 'EVEN Other 2 of 4', style: {} },
      { text: 'ODD Other 3 of 4', style: {} },
      { text: 'EVEN Other 4 of 4', style: {} },
    ])
  })

  it('selects physical pages after assigning footer variants and the full page count', () => {
    const plan = layoutPrintWorkbook([sheet('Report', 2), sheet('Other', 4), sheet('Last', 4)], { fromPage: 2, toPage: 4 })
    expect(plan.totalPages).toBe(5)
    expect(plan.pages.map((page) => [page.sheet, page.position, page.number])).toEqual([
      ['Other', 2, 2], ['Other', 3, 3], ['Last', 4, 4],
    ])
    expect(plan.pages.map((page) => formatPrintHeaderFooter(page.footer!.text, {
      sheetName: page.sheet, pageNumber: page.number, pageCount: plan.totalPages,
    }).center[0])).toEqual([
      { text: 'EVEN Other 2 of 5', style: {} },
      { text: 'ODD Other 3 of 5', style: {} },
      { text: 'EVEN Last 4 of 5', style: {} },
    ])
  })

  it('uses the explicit first number on the first sheet as the next automatic seed', () => {
    expect(layoutPrintWorkbook([sheet('Report', 4, 4), sheet('Other', 4), sheet('Last', 4)]).pages.map((page) => page.number))
      .toEqual([4, 5, 4, 5, 6, 7])
    expect(layoutPrintWorkbook([sheet('Report', 2, -2), sheet('Other', 4), sheet('Last', 4)]).pages.map((page) => page.number))
      .toEqual([-2, -2, -1, 0, 1])
  })

  it('advances the automatic seed after an explicit number on a later sheet', () => {
    const plan = layoutPrintWorkbook([sheet('Report', 2), sheet('Other', 4, 9), sheet('Last', 4)])
    expect(plan.pages.map((page) => page.number)).toEqual([1, 9, 10, 11, 12])
    expect(() => layoutPrintWorkbook([sheet('Report', 2), sheet('Other', 4)], { maxPages: 2 })).toThrow(/page limit/)
  })
})
