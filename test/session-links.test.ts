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
})
