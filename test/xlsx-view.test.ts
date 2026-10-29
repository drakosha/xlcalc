/** Saved sheet views roundtrip without participating in either computation hash. */
import JSZip from 'jszip'
import { afterAll, describe, expect, it } from 'vitest'
import { compute, InputError, type BookInput, type Execution, type FrozenPanes, type SheetView } from '../src/index.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'
afterAll(closeCalcs)
const plain = (): BookInput => ({ sheets: [{ name: 'Hidden', hidden: true, cells: {} }, { name: 'First', cells: { A1: { value: 1 } } }, { name: 'Saved', cells: { C4: { formula: 'First!A1*2' } } }] })
const viewed = (): BookInput => ({ ...plain(), activeSheet: 'Saved', sheets: plain().sheets.map((s) => s.name === 'Saved' ? { ...s, view: { activeCell: 'C4', selection: ['C4:D5', 'F7'], topLeft: 'B3', zoom: 125 } } : s) })

it('reads explicit OOXML fields before the writer supports them', async () => {
  const zip = await JSZip.loadAsync((await writeXlsx({ book: plain() })).data)
  const wb = await zip.file('xl/workbook.xml')!.async('string')
  zip.file('xl/workbook.xml', wb.replace('activeTab="1"', 'activeTab="2"'))
  const path = 'xl/worksheets/sheet3.xml'
  zip.file(path, (await zip.file(path)!.async('string')).replace(/<sheetView[^>]*\/>/, '<sheetView workbookViewId="0" tabSelected="1" topLeftCell="B3" zoomScale="125"><selection activeCell="C4" sqref="C4:D5 F7"/></sheetView>'))
  const read = await readXlsx(await zip.generateAsync({ type: 'uint8array' }))
  expect(read.book.activeSheet).toBe('Saved')
  expect(read.book.sheets[2].view).toEqual(viewed().sheets[2].view)
})

it('writes the saved tab, selection, scroll and zoom, then reads and writes them again', async () => {
  const written = await writeXlsx({ book: viewed(), result: compute(viewed()) })
  const first = await readXlsx(written.data)
  const second = await readXlsx((await writeXlsx({ book: first.book })).data)
  expect([first.book.activeSheet, second.book.activeSheet]).toEqual(['Saved', 'Saved'])
  expect([first.book.sheets[2].view, second.book.sheets[2].view]).toEqual([viewed().sheets[2].view, viewed().sheets[2].view])
  expect([...written.issues, ...first.issues, ...second.issues]).toEqual([])
})

it('keeps the first visible tab as the writer default', async () => {
  const zip = await JSZip.loadAsync((await writeXlsx({ book: plain() })).data)
  expect(await zip.file('xl/workbook.xml')!.async('string')).toContain('activeTab="1"')
  expect(await zip.file('xl/worksheets/sheet2.xml')!.async('string')).toContain('tabSelected="1"')
})

it('rejects malformed view fields and an active hidden or missing sheet', () => {
  for (const activeSheet of ['Missing', 'Hidden', 2]) expect(() => compute({ ...plain(), activeSheet } as never)).toThrow(InputError)
  for (const view of [{ activeCell: 'XFE1' }, { selection: ['XFE1:A1'] }, { zoom: 9 }, { zoom: 401 }, { topLeft: 'B0' }, { tabSelected: 1 }, { unknown: true }])
    expect(() => compute({ sheets: [{ name: 'S', cells: {}, view }] } as never)).toThrow(InputError)
})

describe.each(['closures', 'eval'] as const)('view metadata, %s', (execution: Execution) => {
  it('changes neither hash', () => {
    const a = compute(plain(), { execution })
    const b = compute(viewed(), { execution })
    expect([b.hash, b.inputHash]).toEqual([a.hash, a.inputHash])
  })
  describe.each(KINDS)('%s', (kind) => {
    it('getLayout and snapshots keep the saved view', async () => {
      const s = await (await opener(kind))(viewed(), { execution })
      expect((await s.getLayout()).activeSheet).toBe('Saved')
      expect((await s.getLayout()).sheets[2].view).toEqual(viewed().sheets[2].view)
      expect((await s.getSnapshot()).book.sheets[2].view).toEqual(viewed().sheets[2].view)
      await s.close()
    })
    it('sheet operations keep activeSheet valid', async () => {
      const s = await (await opener(kind))(viewed(), { execution })
      await s.editStructure({ op: 'renameSheet', sheet: 'Saved', name: 'Renamed' })
      expect((await s.getBook()).activeSheet).toBe('Renamed')
      await s.editStructure({ op: 'hideSheet', sheet: 'Renamed' })
      expect((await s.getBook()).activeSheet).toBe('First')
      await s.undo()
      await s.editStructure({ op: 'deleteSheet', sheet: 'Renamed' })
      expect((await s.getBook()).activeSheet).toBe('First')
      await s.close()
    })
  })
})

it('selects the active tab alone, and keeps a hidden preceding sheet and the selection of the active frozen pane', async () => {
  const b = viewed()
  b.sheets[2].frozen = { rows: 1, columns: 1 }
  b.sheets[2].view!.topLeft = 'A1'
  const first = await readXlsx((await writeXlsx({ book: b })).data)
  // another sheet made active after reading: its tab alone is selected, the book does not open grouped
  const data = (await writeXlsx({ book: { ...first.book, activeSheet: first.book.sheets[1].name } })).data
  const zip = await JSZip.loadAsync(data)
  const selected = await Promise.all(first.book.sheets.map(async (_, i) => (await zip.file(`xl/worksheets/sheet${i + 1}.xml`)!.async('string')).includes('tabSelected="1"')))
  expect(selected).toEqual(first.book.sheets.map((_, i) => i === 1))
  const second = await readXlsx((await writeXlsx({ book: first.book })).data)
  expect(second.book.sheets[1].view).toBeUndefined()
  expect(second.book.sheets[2].view).toEqual(b.sheets[2].view)
  expect(second.book.sheets[2].frozen).toEqual({ rows: 1, columns: 1 })
})

it("reads the selection of the active pane, the top left one where the pane names none, and no other pane's", async () => {
  const zip = await JSZip.loadAsync((await writeXlsx({ book: { sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] } })).data)
  const path = 'xl/worksheets/sheet1.xml'
  const xml = await zip.file(path)!.async('string')
  const read = async (body: string) => {
    zip.file(path, xml.replace(/<sheetView[^>]*\/>/, `<sheetView workbookViewId="0" tabSelected="1"><pane xSplit="3640" ySplit="1940" topLeftCell="D6"/>${body}</sheetView>`))
    return (await readXlsx(await zip.generateAsync({ type: 'uint8array' }))).book.sheets[0].view
  }
  // as Excel for Mac saves a window split with its cursor on A1: the other panes' selections, none of the first
  const others = '<selection pane="bottomLeft" activeCell="A6" sqref="A6"/><selection pane="topRight" activeCell="D1" sqref="D1"/><selection pane="bottomRight" activeCell="D6" sqref="D6"/>'
  expect((await read(others))?.activeCell).toBeUndefined()
  expect((await read(`<selection activeCell="B2" sqref="B2"/>${others}`))?.activeCell).toBe('B2')
})

it('roundtrips an explicitly saved first active sheet', async () => {
  const book: BookInput = { activeSheet: 'S', sheets: [{ name: 'S', cells: {} }, { name: 'Other', cells: {} }] }
  const data = (await writeXlsx({ book })).data
  expect((await readXlsx(data)).book.activeSheet).toBe('S')
})

describe.each(['closures', 'eval'] as const)('frozen panes with a saved scroll position (fix-u325), %s', (execution: Execution) => {
  const book = (frozen: FrozenPanes, topLeft: string): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: 1 } }, frozen, view: { topLeft } }] })
  const cycle = async (b: BookInput) => {
    const r = await readXlsx((await writeXlsx({ book: b })).data)
    return { book: r.book, issues: r.issues }
  }

  it.each([
    [{ rows: 2, columns: 1 }, 'C5'],
    [{ rows: 1 }, 'A3'],
    [{ columns: 2 }, 'F1'],
    [{ rows: 1 }, 'C40'],
    [{ rows: 3, columns: 2 }, 'A1'],
  ] as [FrozenPanes, string][])('%j over %s round-trips, again and again, and keeps inputHash', async (frozen, topLeft) => {
    const b = book(frozen, topLeft)
    let r = await cycle(b)
    const view = r.book.sheets[0].view
    for (let i = 0; i < 3; i++) {
      expect([r.book.sheets[0].frozen, r.book.sheets[0].view?.topLeft, r.issues]).toEqual([frozen, topLeft, []])
      expect(r.book.sheets[0].view).toEqual(view)
      r = await cycle(r.book)
    }
    expect(compute(r.book, { execution }).inputHash).toBe(compute(b, { execution }).inputHash)
  })

  it('the scroll goes to the pane after the split, the view starts at the first row and column', async () => {
    const zip = await JSZip.loadAsync((await writeXlsx({ book: book({ rows: 2, columns: 1 }, 'C5') })).data)
    const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(xml).toContain('<sheetView workbookViewId="0" tabSelected="1" topLeftCell="A1"><pane xSplit="1" ySplit="2" topLeftCell="C5" activePane="bottomRight" state="frozen"/>')
  })

  it('a file frozen over a scrolled sheet is read once from the first row, and written so', async () => {
    const zip = await JSZip.loadAsync((await writeXlsx({ book: { sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] } })).data)
    const path = 'xl/worksheets/sheet1.xml'
    zip.file(path, (await zip.file(path)!.async('string')).replace(/<sheetView[^>]*\/>/, '<sheetView workbookViewId="0" topLeftCell="A5"><pane ySplit="2" topLeftCell="A7" activePane="bottomLeft" state="frozen"/></sheetView>'))
    const first = await readXlsx(await zip.generateAsync({ type: 'uint8array' }))
    expect(first.book.sheets[0].frozen).toEqual({ rows: 6 })
    expect(first.issues.map((i) => i.code)).toEqual(['frozenPanes'])
    const second = await cycle(first.book)
    const third = await cycle(second.book)
    expect([second.book.sheets[0].frozen, third.book.sheets[0].frozen, second.issues, third.issues]).toEqual([{ rows: 6 }, { rows: 6 }, [], []])
    expect(third.book.sheets[0].view).toEqual(second.book.sheets[0].view)
  })
})

describe('a window split into panes', () => {
  const split = (view: SheetView, frozen?: FrozenPanes): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: 1 } }, view, ...(frozen ? { frozen } : {}) }] })
  const sheetViews = async (data: Uint8Array) => (await (await JSZip.loadAsync(data)).file('xl/worksheets/sheet1.xml')!.async('string')).match(/<sheetViews>.*<\/sheetViews>/)![0]

  it('is checked, and changes neither hash', () => {
    const view = { split: { x: 120, y: 60, topLeft: 'F20', activePane: 'bottomRight' as const } }
    expect([compute(split(view)).hash, compute(split(view)).inputHash]).toEqual([compute(split({})).hash, compute(split({})).inputHash])
    const refused: [unknown, string, string][] = [
      [{}, 'book.sheets[0].view.split', 'expected a line: x or y above 0'],
      [{ x: 0 }, 'book.sheets[0].view.split', 'expected a line: x or y above 0'],
      [{ x: -1 }, 'book.sheets[0].view.split.x', 'expected a number from 0 to 100000, got number -1'],
      [{ y: 10, activePane: 'topRight' }, 'book.sheets[0].view.split.activePane', 'expected one of "topLeft", "bottomLeft", got the string "topRight"'],
      [{ x: 10, topLeft: 'A0' }, 'book.sheets[0].view.split.topLeft', ''],
      [{ x: 10, state: 'split' }, 'book.sheets[0].view.split', ''],
    ]
    for (const [v, path, message] of refused) {
      const e = (() => { try { compute(split({ split: v as never })) } catch (x) { return x as InputError } })()
      expect(e, JSON.stringify(v)).toBeInstanceOf(InputError)
      expect(e!.path).toBe(path)
      if (message) expect(e!.message).toContain(message)
    }
  })

  it('is written as Excel writes it, in twentieths of a point, and read back', async () => {
    const view: SheetView = { split: { x: 120, y: 60.25, topLeft: 'F20', activePane: 'bottomRight' }, activeCell: 'G21', selection: ['G21'] }
    const data = (await writeXlsx({ book: split(view) })).data
    expect(await sheetViews(data)).toBe('<sheetViews><sheetView workbookViewId="0" tabSelected="1"><pane xSplit="2400" ySplit="1205" topLeftCell="F20" activePane="bottomRight"/><selection pane="bottomRight" activeCell="G21" sqref="G21"/></sheetView></sheetViews>')
    const read = await readXlsx(data)
    expect([read.book.sheets[0].view, read.book.sheets[0].frozen, read.issues]).toEqual([view, undefined, []])
    // the first pane active: no pane named, as Excel writes it
    expect(await sheetViews((await writeXlsx({ book: split({ split: { y: 30 }, activeCell: 'B2', selection: ['B2'] }) })).data)).toBe('<sheetViews><sheetView workbookViewId="0" tabSelected="1"><pane ySplit="600"/><selection activeCell="B2" sqref="B2"/></sheetView></sheetViews>')
  })

  it('frozen panes take its place; unfrozen, the split is written again', async () => {
    const view: SheetView = { split: { x: 120, topLeft: 'F1' } }
    const frozen = await readXlsx((await writeXlsx({ book: split(view, { rows: 2 }) })).data)
    expect([frozen.book.sheets[0].frozen, frozen.book.sheets[0].view]).toEqual([{ rows: 2 }, undefined])
    for (const kind of KINDS) {
      const s = await (await opener(kind))(split(view), {})
      await s.setLayout([{ sheet: 'S', frozen: { rows: 2 } }])
      await s.setLayout([{ sheet: 'S', frozen: null }])
      expect((await readXlsx((await writeXlsx({ book: await s.getBook() })).data)).book.sheets[0].view, kind).toEqual(view)
      await s.close()
    }
  })

  it("writeKept puts the workbook's pane in place of the file's and keeps the rest of the view", async () => {
    const zip = await JSZip.loadAsync((await writeXlsx({ book: split({}) })).data)
    const path = 'xl/worksheets/sheet1.xml'
    // as Excel for Mac saved a split window, with a pivot table's selection and an extension of its own
    const excel = '<sheetView tabSelected="1" workbookViewId="0"><pane xSplit="3640" ySplit="1940" topLeftCell="D6"/><selection pane="bottomLeft" activeCell="A6" sqref="A6"/><selection pane="topRight" activeCell="D1" sqref="D1"/><selection pane="bottomRight" activeCell="D6" sqref="D6"/><pivotSelection pane="bottomRight"><pivotArea/></pivotSelection><extLst><ext uri="{X}"/></extLst></sheetView>'
    zip.file(path, (await zip.file(path)!.async('string')).replace(/<sheetView[^>]*\/>/, excel))
    const data = await zip.generateAsync({ type: 'uint8array' })
    const read = await openKept(data)
    expect(read.book.sheets[0].view).toEqual({ split: { x: 182, y: 97, topLeft: 'D6' } })
    expect(await sheetViews((await writeKept(data, { book: read.book })).data)).toBe(`<sheetViews>${excel}</sheetViews>`)
    const moved = { ...read.book, sheets: [{ ...read.book.sheets[0], view: { split: { x: 200, topLeft: 'H1', activePane: 'topRight' as const } } }] }
    expect(await sheetViews((await writeKept(data, { book: moved })).data)).toBe('<sheetViews><sheetView workbookViewId="0" tabSelected="1"><pane xSplit="4000" topLeftCell="H1" activePane="topRight"/><pivotSelection pane="bottomRight"><pivotArea/></pivotSelection><extLst><ext uri="{X}"/></extLst></sheetView></sheetViews>')
  })
})
