import { describe, expect, it } from 'vitest'
import { layoutPrintPages, layoutPrintWorkbook, printSheetGeometry, type PrintNotes, type PrintWorkbookSheet } from '../src/xlsx-print.js'

function sheet(notes: PrintNotes = { cells: {} }): PrintWorkbookSheet {
  return {
    settings: {
      sheet: 'Report',
      setup: { cellComments: 'atEnd' },
      margins: { left: 0, right: 0, top: 0, bottom: 0, header: 0, footer: 0 },
      headerFooter: { differentFirst: true, differentOddEven: true, firstFooter: 'First', oddFooter: 'Odd', evenFooter: 'Even' },
    },
    geometry: printSheetGeometry({ rows: { '4': { hidden: true } } }, { usedRange: 'A1:C4' }),
    target: { width: 100, height: 60 },
    notes: { metrics: () => ({ lineHeight: 10, measureText: (text) => text.length * 5 }), ...notes },
  }
}

describe('printing worksheet notes', () => {
  it('prints notes in area order, including hidden cells and duplicates in overlaps', () => {
    const report = sheet({ cells: {
      E1: { note: { text: 'Outside' } },
      A2: { note: { text: 'One' } },
      C1: { note: { text: 'Two', author: 'Author' } },
      B4: { note: { text: 'Three' } },
      B2: { note: { text: '' } },
      A1: { note: { text: 'Only in titles' } },
    } })
    report.target.width = 120
    report.settings.areas = ['B1:C4', 'A2:B2']
    report.settings.repeatRows = { first: 1, last: 1 }
    const plan = layoutPrintWorkbook([report])
    const notes = plan.pages.flatMap((page) => page.noteLines ?? [])
    expect(notes.filter((line) => line.kind === 'address').map((line) => line.cell)).toEqual(['C1', 'B2', 'B4', 'A2', 'B2'])
    expect(notes.filter((line) => line.kind === 'body').map((line) => line.text)).toEqual(['Two', 'Three', 'One'])
    expect(plan.pages.slice(0, 2).map((page) => [page.area, page.noteLines])).toEqual([[1, undefined], [2, undefined]])
    expect(plan.pages.slice(2).every((page) => page.area === 0 && page.regions.length === 0)).toBe(true)
  })

  it('can end a page at a note heading and carries over text and numbering', () => {
    const report = sheet({ cells: {
      A1: { note: { text: 'short' } },
      A2: { note: { text: 'short' } },
      A3: { note: { text: 'first words here\n\nnext line' } },
    } })
    report.settings.areas = ['A1:A3']
    const last = sheet()
    last.settings = { ...last.settings, sheet: 'Last', areas: ['A1'] }
    const plan = layoutPrintWorkbook([report, last], { fromPage: 2, toPage: 4 })
    expect(plan.totalPages).toBe(4)
    expect(plan.pages.map((page) => [page.sheet, page.position, page.number, page.footer?.text])).toEqual([
      ['Report', 2, 2, 'Even'], ['Report', 3, 3, 'Odd'], ['Last', 4, 4, 'Even'],
    ])
    expect(plan.pages[0].noteLines?.map((line) => [line.text, line.bounds.y])).toEqual([
      ['A1', 0], ['short', 10], ['A2', 30], ['short', 40], ['A3', 60],
    ])
    expect(plan.pages[1].noteLines?.map((line) => [line.text, line.bounds.y])).toEqual([
      ['first words', 0], ['here', 10], ['', 20], ['next line', 30],
    ])
    expect(plan.pages[1].noteLines?.[0].label?.text).toBe('Note:')
    expect(() => layoutPrintWorkbook([report], { maxPages: 2 })).toThrow(/page limit/)
  })

  it('retains a blank grid page before the notes when the entire area is hidden', () => {
    const report = sheet({ cells: { B4: { note: { text: 'Hidden' } } } })
    report.settings.areas = ['B4']
    const { settings, geometry, target, notes } = report
    const plan = layoutPrintPages(settings, geometry, target, notes)
    expect(plan.pages.map((page) => [page.area, page.regions.length, page.noteLines?.map((line) => line.text)])).toEqual([
      [1, 0, undefined], [0, 0, ['B4', 'Hidden']],
    ])
    expect(layoutPrintPages({ ...settings, setup: { cellComments: 'none' } }, geometry, target, notes).pages).toEqual([])
  })
})
