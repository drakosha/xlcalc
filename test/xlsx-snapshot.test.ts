/**
 * `writeXlsx` takes a snapshot `{ book, result? }` and which values of formulas to store (`values`): `require` (the
 * default with a result: every formula and spilled cell has its value, else the result is not the book's), `omit`
 * (the default without a result, reported when the book has formulas), `preserve` (the values the result has, the
 * rest reported); the old form `writeXlsx(book, { result })` writes the same bytes. `readXlsx` gives `cached` in the
 * shape of a result, which writes the file again as the result of its book. Wrong arguments are `InputError`s with
 * what to pass instead.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { BookInput } from '../src/contract.js'
import { compute, InputError, openSession } from '../src/index.js'
import { readXlsx, writeXlsx, type XlsxCached } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { codes, storedOf, storedValues } from './xlsx-fixtures.js'

const clock = { now: '2026-10-10T12:00:00Z', seed: 3 }
const part = async (data: Uint8Array, name: string) => (await openArchive(data).text(name))!

const book: BookInput = {
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 120 }, A2: { value: 80 },
        B1: { formula: 'SUM(A1:A2)' },
        C1: { formula: 'SEQUENCE(2)' },
        D1: { formula: 'A1:A2*10', array: 'D1:D2' },
        E1: { formula: 'SEQUENCE(1)' },
      },
    },
    { name: 'T', cells: { A1: { value: 'x' } } },
  ],
}

/** The result without some of its cells: of another book, or cut. */
function without(cells: string[]): XlsxCached {
  const r = compute(book, { clock })
  return { sheets: r.sheets.map((s, i) => ({ name: s.name, cells: Object.fromEntries(Object.entries(s.cells).filter(([a]) => i > 0 || !cells.includes(a))) })) }
}

describe('writeXlsx with a snapshot', () => {
  it('writes what the old form wrote: the same bytes, the same issues', async () => {
    const result = compute(book, { clock })
    const now = await writeXlsx({ book, result })
    expect(now.data).toEqual((await writeXlsx(book, { result })).data)
    expect(now.issues).toEqual([])
    expect((await writeXlsx({ book, result }, { mode: 'values' })).data).toEqual((await writeXlsx(book, { result, mode: 'values' })).data)
  })

  it("require, the default with a result: every formula and every cell of an array has its value, else it is not the book's result", async () => {
    await expect(writeXlsx({ book, result: without(['B1']) })).rejects.toThrow(new InputError('snapshot.result', `expected the book's result: the formula of 'S'!B1 has no value in the result (values: "preserve" writes the values it has)`))
    await expect(writeXlsx({ book, result: without(['C2']) })).rejects.toThrow(/'S'!C2 of the array of C1 has no value in the result/)
    await expect(writeXlsx({ book, result: without(['D2']) }, { values: 'require' })).rejects.toThrow(/'S'!D2 of the array of D1 has no value/)
    await expect(writeXlsx(book, { result: without(['B1']) })).rejects.toThrow(new InputError('opts.result', `expected the book's result: the formula of 'S'!B1 has no value in the result (values: "preserve" writes the values it has)`))
  })

  it('preserve: the values the result has; a formula without one is written without it, and reported', async () => {
    const { data, issues } = await writeXlsx({ book, result: without(['B1', 'E1']) }, { values: 'preserve' })
    expect(codes(issues)).toEqual(['noCachedValues B1 E1'])
    expect(issues[0].message).toBe('a formula without a value in the result is written without one: Excel computes it when it opens the file')
    expect(await part(data, 'xl/worksheets/sheet1.xml')).toContain('<c r="B1" cm="1"><f t="array" ref="B1">SUM(A1:A2)</f></c>')
    expect(storedValues(await readXlsx(data))).toEqual({ C1: 1, C2: 2, D1: 1200, D2: 800 })
    const values = await writeXlsx({ book, result: without(['B1']) }, { mode: 'values', values: 'preserve' })
    expect(codes(values.issues)).toEqual(['noCachedValues B1'])
    expect((await readXlsx(values.data)).book.sheets[0].cells.B1).toBeUndefined()
  })

  it('omit, the default without a result: no values of formulas, reported when the book has formulas; asked for, not reported', async () => {
    const plain = await writeXlsx({ book })
    expect(plain.issues).toEqual([{ code: 'noCachedValues', message: 'no result was given: the file stores no values of formulas, and Excel computes them when it opens the file' }])
    expect(await part(plain.data, 'xl/workbook.xml')).toContain('<calcPr fullCalcOnLoad="1"/>')
    expect((await writeXlsx(book)).issues).toEqual(plain.issues)
    expect((await writeXlsx({ book }, { values: 'omit' })).issues).toEqual([])
    expect((await writeXlsx({ book: { sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] } })).issues).toEqual([])
    // with a result too: the file is the one written without it
    const omitted = await writeXlsx({ book, result: compute(book, { clock }) }, { values: 'omit' })
    expect(omitted.data).toEqual((await writeXlsx({ book }, { values: 'omit' })).data)
    expect(storedValues(await readXlsx(omitted.data))).toEqual({})
  })

  it('every file with formulas asks Excel to compute them when it opens it, whatever values it stores', async () => {
    for (const values of ['require', 'preserve', 'omit'] as const) expect(await part((await writeXlsx({ book, result: compute(book, { clock }) }, { values })).data, 'xl/workbook.xml')).toContain('fullCalcOnLoad="1"')
  })

  it('refuses what it cannot take, saying what to pass instead', async () => {
    const session = openSession(book, { clock })
    await expect(writeXlsx(session as never)).rejects.toThrow(new InputError('snapshot', 'a session is not a snapshot: pass session.getSnapshot() (awaited for an asynchronous session)'))
    await expect(writeXlsx(book, { result: session as never })).rejects.toThrow(new InputError('opts.result', 'a session is not a snapshot: pass session.getSnapshot() (awaited for an asynchronous session)'))
    await expect(writeXlsx({ book, result: session as never })).rejects.toThrow(InputError)
    session.close()
    const result = "a result: compute's, a session's getSnapshot().result or readXlsx's cached, { sheets: [{ name, cells }] }"
    await expect(writeXlsx({ book, result: [{ name: 'S', cells: { B1: 200 } }] as never })).rejects.toThrow(new InputError('snapshot.result', `expected ${result}, got an array`))
    await expect(writeXlsx({ book, result: { cells: {} } as never })).rejects.toThrow(new InputError('snapshot.result', `expected ${result}, got an object without sheets`))
    await expect(writeXlsx({ book, result: { sheets: [{ name: 'S', cells: { B1: 200 } }, { name: 'T', cells: {} }] } as never })).rejects.toThrow(new InputError('snapshot.result.sheets[0].cells["B1"]', 'expected a cell of a result, such as { value: 200 }, got number 200'))
    await expect(writeXlsx({ book, result: compute({ sheets: [{ name: 'X', cells: {} }] }) })).rejects.toThrow(/its sheets differ from the book's/)
    await expect(writeXlsx({ book, result: compute(book) }, { result: compute(book) } as never)).rejects.toThrow(/the result goes in the snapshot: writeXlsx\(\{ book, result \}\)/)
    const read = await readXlsx((await writeXlsx({ book, result: compute(book) })).data)
    await expect(writeXlsx(read as never)).rejects.toThrow(/a file read: pass \{ book: read\.book, result: read\.cached \}/)
    await expect(writeXlsx({ book }, { values: 'require' })).rejects.toThrow(new InputError('snapshot.result', 'values: "require" takes the values of formulas from the result: pass { book, result }'))
    await expect(writeXlsx({ book, result: compute(book) }, { mode: 'values', values: 'omit' })).rejects.toThrow(/the values mode writes values only/)
    await expect(writeXlsx({ book }, { mode: 'values' })).rejects.toThrow(new InputError('snapshot.result', 'the values mode takes the values from the result: pass { book, result }'))
    await expect(writeXlsx({ book }, { values: 'all' as never })).rejects.toThrow(/expected "require", "omit" or "preserve"/)
    await expect(writeXlsx({ book: { sheets: [{ name: 'a:b', cells: {} }] } })).rejects.toThrow(InputError)
    await expect(writeXlsx(undefined as never)).rejects.toThrow(new InputError('snapshot', 'expected a snapshot { book, result }, got undefined'))
  })
})

describe('readXlsx', () => {
  it("gives cached in the shape of a result: a formula's value with its array's range, each cell of the array with its formula's address", async () => {
    const result = compute(book, { clock })
    const read = await readXlsx((await writeXlsx({ book, result })).data)
    expect(read.cached).toEqual(storedOf(result))
    expect(read.cached.sheets[0].cells).toEqual({
      B1: { value: 200 },
      C1: { value: 1, spill: 'C1:C2' }, C2: { value: 2, spillOf: 'C1' },
      D1: { value: 1200, spill: 'D1:D2' }, D2: { value: 800, spillOf: 'D1' },
      E1: { value: 1 },
    })
    expect(read.cached.sheets[0].cells.B1.value).toBe(result.sheets[0].cells.B1.value)
  })

  it("cached writes the file again as the result of its book: Excel's values stay, the bytes too", async () => {
    const first = await writeXlsx({ book, result: compute(book, { clock }) })
    const read = await readXlsx(first.data)
    const again = await writeXlsx({ book: read.book, result: read.cached })
    expect(again.issues).toEqual([])
    expect(again.data).toEqual(first.data)
    // a file Excel saved: its values written back as they were
    const saved = await readXlsx(readFileSync(new URL('./excel/xlsx/roundtrip/spills.xlsx', import.meta.url)))
    const back = await readXlsx((await writeXlsx({ book: saved.book, result: saved.cached })).data)
    expect(back.cached).toEqual(saved.cached)
  })

  it("takes the file's bytes as a Uint8Array, a Buffer, an ArrayBuffer or another view of them; a path is refused with what to pass", async () => {
    const { data } = await writeXlsx({ book, result: compute(book, { clock }) })
    const expected = (await readXlsx(data)).book
    const copy = new ArrayBuffer(data.byteLength + 8)
    new Uint8Array(copy, 8).set(data)
    for (const bytes of [Buffer.from(data), data.slice().buffer, new DataView(copy, 8), new Uint8Array(copy, 8)]) expect((await readXlsx(bytes as Uint8Array)).book).toEqual(expected)
    await expect(readXlsx('book.xlsx' as never)).rejects.toThrow(new InputError('data', `expected the file's bytes, got the string "book.xlsx": pass the file's bytes: Uint8Array, ArrayBuffer or Buffer, e.g. await readFile(path)`))
    await expect(readXlsx(new Blob([data]) as never)).rejects.toThrow(/e\.g\. await blob\.arrayBuffer\(\)/)
    await expect(readXlsx(undefined as never)).rejects.toThrow(InputError)
  })
})
