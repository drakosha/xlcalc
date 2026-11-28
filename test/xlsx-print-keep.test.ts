import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { createSession } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { printingSnapshot, worksheetPrintOf } from './print-fixtures.js'
import type { SheetPrint } from '../src/xlsx-print.js'
import { changedParts, partsOf, withParts } from './keep-files.js'

describe('xlsx/print/keep', () => {
  it('writes print settings in place and preserves the sheet’s drawing and printer relationship reference', async () => {
    const original = await withParts({
      sheets: [{ name: 'Data', tabColor: '#AA6600', outline: { summaryBelow: false }, cells: { A1: { value: 1 }, A2: { value: 2 } } }],
    }, {
      chart: { sheet: 0, series: ["'Data'!$A$1:$A$2"], anchor: [1, 0, 8, 3] },
      printArea: { sheet: 0, ref: "'Data'!$A$1:$A$5" },
      edit: { 'xl/worksheets/sheet1.xml': (xml) => xml.replace('<drawing', '<pageSetup orientation="portrait" r:id="rId77"/><drawing') },
    })
    const opened = await openKept(original)
    expect(worksheetPrintOf(opened.book)[0]).toMatchObject({ sheet: 'Data', areas: ['A1:A5'] })
    const snapshot = { book: opened.book, result: opened.cached }
    const print: SheetPrint = {
      sheet: 'Data', areas: ['A1:C12'], repeatRows: { first: 1, last: 1 },
      setup: { fitToPage: true, fitToWidth: 1, fitToHeight: 0, orientation: 'landscape' },
      margins: { left: 0.6, right: 0.6, top: 0.7, bottom: 0.7, header: 0.3, footer: 0.3 },
      headerFooter: { oddFooter: '&RPage &P of &N' },
    }
    expect((await writeKept(original, snapshot)).data).toEqual(original)
    const { data } = await writeKept(original, printingSnapshot(snapshot, [print]))
    expect(await changedParts(original, data)).toEqual(['xl/workbook.xml', 'xl/worksheets/sheet1.xml'])
    const sheetXml = (await partsOf(data)).get('xl/worksheets/sheet1.xml')!
    expect(sheetXml).toMatch(/<tabColor[^>]*\/><outlinePr[^>]*\/><pageSetUpPr[^>]*\/><\/sheetPr>/)
    expect(sheetXml).toMatch(/<pageMargins[^>]*\/><pageSetup r:id="rId77"[^>]*\/><headerFooter[\s\S]*<\/headerFooter><drawing/)
    const reread = await openKept(data)
    expect(worksheetPrintOf(reread.book)).toEqual([print])
    expect(reread.printIssues).toContainEqual({ sheet: 'Data', part: 'pageSetup', reason: 'printerSettings' })

    const cleared = await writeKept(data, printingSnapshot({ book: reread.book, result: reread.cached }, [{ sheet: 'Data' }]))
    const left = await openKept(cleared.data)
    expect(worksheetPrintOf(left.book)).toEqual([])
    expect((await partsOf(cleared.data)).get('xl/worksheets/sheet1.xml')).toContain('<pageSetup r:id="rId77"/>')
  })

  it('keeps print settings in the session through renaming, layout edits, undo and both XLSX writers', async () => {
    const original = await withParts({ sheets: [{ name: 'Data', cells: { B2: { value: 1 } } }] }, {
      printArea: { sheet: 0, ref: "'Data'!$B$2:$D$8" },
    })
    const opened = await openKept(original)
    const session = createSession(opened.book)
    expect(session.editStructure({ op: 'insertRows', sheet: 'Data', row: 1 }).ok).toBe(true)
    expect(session.editStructure({ op: 'renameSheet', sheet: 'Data', name: 'Report' }).ok).toBe(true)
    expect(session.getSnapshot().book.sheets[0].print?.areas).toEqual(['B3:D9'])
    const print = { areas: ['B3:D9'], rowBreaks: [{ before: 5, manual: true }], setup: { orientation: 'landscape' as const } }
    session.setLayout([{ sheet: 'Report', print }])
    expect(session.getSnapshot().book.sheets[0].print).toEqual(print)
    session.undo()
    expect(session.getSnapshot().book.sheets[0].print).toEqual({ areas: ['B3:D9'] })
    session.redo()
    const snapshot = session.getSnapshot()
    for (const { data } of [await writeXlsx(snapshot), await writeKept(original, snapshot)]) {
      expect((await readXlsx(data)).book.sheets[0]).toMatchObject({ name: 'Report', print })
    }
    session.close()
  })

  it('scopes print titles by the worksheet’s actual position, including chart sheets', async () => {
    const original = readFileSync(new URL('./excel/keep/open/chartsheet-parts.xlsx', import.meta.url))
    const read = await openKept(original)
    expect(read.book.sheets[0].name).toBe('Data')
    const settings: SheetPrint = { sheet: 'Data', repeatRows: { first: 2, last: 3 }, areas: ['B2:D9'] }
    const { data } = await writeKept(original, printingSnapshot({ book: read.book, result: read.cached }, [settings]))
    const parts = await partsOf(data)
    const workbook = parts.get('xl/workbook.xml')!
    expect(workbook).toMatch(/name="_xlnm.Print_Titles" localSheetId="1"/)
    expect(workbook.indexOf('<definedNames>')).toBeGreaterThan(workbook.indexOf('</sheets>'))
    expect(workbook.indexOf('<definedNames>')).toBeLessThan(workbook.indexOf('<calcPr'))
    expect(worksheetPrintOf((await openKept(data)).book).find((sheet) => sheet.sheet === 'Data')).toEqual(settings)
    const before = await partsOf(original)
    for (const [name, xml] of before) if (name.startsWith('xl/charts')) expect(parts.get(name)).toBe(xml)
  })
})
