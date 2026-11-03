/**
 * `Session.setExternalBooks`, in every kind of session: a host adds a link with the values it read and extends one,
 * formulas typed over them find them by file name and read the values, the workbook is calculated again in one step of
 * undo, and each kind gives the same workbook and hash; the links are checked as a workbook's are.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { InputError, type BookInput, type ExternalBook } from '../src/index.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const clock = { now: '2026-10-27T12:00:00Z', seed: 1 }
const book = (): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { formula: "'[1]Data'!B2*2" } } }], externalBooks: [{ target: 'Other.xlsx', sheets: [{ name: 'Data', cells: { B2: 5 } }] }] })
const links: ExternalBook[] = [
  { target: 'Other.xlsx', sheets: [{ name: 'Data', cells: { B2: 6, C3: 7 } }] },
  { target: 'Отчёт.xlsx', sheets: [{ name: 'Лист', cells: { I8: 42 } }] },
]

describe.each(KINDS)('setExternalBooks in %s', (kind) => {
  it('new values and a new link are read by the formulas over them, in one step of undo, as in each kind', async () => {
    const open = await opener(kind)
    const s = await open(book(), { clock })
    const r = await s.setExternalBooks(links)
    expect(r.changes).toEqual([expect.objectContaining({ sheet: 'S', cell: 'A1', value: 12 })])
    const sync = await (await opener('a synchronous session'))(book(), { clock })
    await sync.setExternalBooks(links)
    expect([await s.getBook(), await s.getHash()]).toEqual([await sync.getBook(), await sync.getHash()])
    // a formula typed over the links finds them by file name (and marks them, as Excel does)
    expect(await s.setCellText('S', 'A2', "='[Отчёт.xlsx]Лист'!$I$8+[Other.xlsx]Data!C3")).toMatchObject({ ok: true })
    expect(await s.getCell('S', 'A2')).toMatchObject({ formula: "'[2]Лист'!$I$8+'[1]Data'!C3", value: 49 })
    expect(s.history).toEqual({ undo: 2, redo: 0 })
    await s.undo()
    await s.undo()
    expect((await s.getBook()).externalBooks?.map((b) => b.sheets[0].cells)).toEqual([{ B2: 5 }])
    expect(await s.getCell('S', 'A1')).toMatchObject({ value: 10 })
    await expect(s.setExternalBooks([{ sheets: [] } as never])).rejects.toBeInstanceOf(InputError)
    await Promise.all([s.close(), sync.close()])
  })

  it('a link the host holds open: a formula typed over it marks nothing and a cell it does not hold is blank; closed again, typing marks it', async () => {
    const open = await opener(kind)
    const s = await open(book(), { clock })
    const held: ExternalBook[] = [{ target: 'Other.xlsx', sheets: [{ name: 'Data', cells: { B2: 5, H1: 10, H2: 20 } }], open: true }]
    await s.setExternalBooks(held)
    expect(await s.setCellText('S', 'A2', '=SUM([Other.xlsx]Data!H:H)+[Other.xlsx]Data!Z99')).toMatchObject({ ok: true })
    expect(await s.setCellText('S', 'A3', '=ISBLANK([Other.xlsx]Data!Z99)')).toMatchObject({ ok: true })
    expect([(await s.getCell('S', 'A2'))?.value, (await s.getCell('S', 'A3'))?.value]).toEqual([30, true])
    expect((await s.getBook()).externalBooks).toEqual(held)
    // closed: the values stay, unmarked, until a formula over the link is typed
    await s.setExternalBooks([{ target: 'Other.xlsx', sheets: held[0].sheets }])
    expect((await s.getCell('S', 'A2'))?.value).toBe(30)
    await s.setCellText('S', 'A4', '=[Other.xlsx]Data!B2')
    expect([(await s.getCell('S', 'A2'))?.value, (await s.getCell('S', 'A3'))?.value]).toEqual([{ error: 'REF' }, false])
    expect((await s.getBook()).externalBooks![0].sheets[0].refreshError).toBe(true)
    await s.undo()
    await s.undo()
    expect((await s.getBook()).externalBooks).toEqual(held)
    await expect(s.setExternalBooks([{ ...held[0], open: false } as never])).rejects.toThrow('books[0].open: expected true, got boolean false')
    await expect(s.setExternalBooks([{ ...held[0], sheets: [{ name: 'Data', refreshError: true }] }])).rejects.toThrow('a linked workbook held open is read: none of its sheets has refreshError')
    await s.close()
  })
})
