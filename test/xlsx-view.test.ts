/** Saved sheet views roundtrip without participating in either computation hash. */
import JSZip from 'jszip'
import { afterAll, describe, expect, it } from 'vitest'
import { compute, InputError, type BookInput, type Execution } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'
afterAll(closeCalcs)
const plain = (): BookInput => ({ sheets: [{ name: 'Hidden', hidden: true, cells: {} }, { name: 'First', cells: { A1: { value: 1 } } }, { name: 'Saved', cells: { C4: { formula: 'First!A1*2' } } }] })
const viewed = (): BookInput => ({ ...plain(), activeSheet: 'Saved', sheets: plain().sheets.map((s) => s.name === 'Saved' ? { ...s, view: { activeCell: 'C4', selection: ['C4:D5', 'F7'], topLeft: 'B3', zoom: 125, tabSelected: true } } : s) })

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

it('keeps grouped tabs, a hidden preceding sheet and the selection of the active frozen pane', async () => {
  const b = viewed()
  b.sheets[1].view = { tabSelected: true }
  b.sheets[2].frozen = { rows: 1, columns: 1 }
  b.sheets[2].view!.topLeft = 'A1'
  const first = await readXlsx((await writeXlsx({ book: b })).data)
  const second = await readXlsx((await writeXlsx({ book: first.book })).data)
  expect(second.book.sheets[1].view).toEqual({ tabSelected: true })
  expect(second.book.sheets[2].view).toEqual(b.sheets[2].view)
  expect(second.book.sheets[2].frozen).toEqual({ rows: 1, columns: 1 })
})

it('roundtrips an explicitly saved first active sheet', async () => {
  const book: BookInput = { activeSheet: 'S', sheets: [{ name: 'S', cells: {} }, { name: 'Other', cells: {} }] }
  const data = (await writeXlsx({ book })).data
  expect((await readXlsx(data)).book.activeSheet).toBe('S')
})
