/**
 * A sheet's tab colour and its view's gridlines and headings as input, the tab colour in every kind of session (a
 * layout edit, one step of undo), and both through `writeXlsx`, `readXlsx` and `writeKept`.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, InputError, type BookInput } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const book = (): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: 1 } }, tabColor: '#FF0000', view: { gridlines: false } }] })

describe('as input', () => {
  it('takes a colour and booleans, and refuses others', () => {
    expect(compute(book()).hash).toBe(compute({ sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] }).hash)
    expect(() => compute({ sheets: [{ name: 'S', cells: {}, tabColor: 'red' }] })).toThrow(InputError)
    expect(() => compute({ sheets: [{ name: 'S', cells: {}, view: { gridlines: 0 as never } }] })).toThrow(new InputError('book.sheets[0].view.gridlines', 'expected a boolean, got number 0'))
  })
})

describe.each(KINDS)('%s', (kind) => {
  it('sets, changes and removes the tab colour by a layout edit, one step of undo each; keeps the view', async () => {
    const s = await (await opener(kind))(book(), {})
    await s.setLayout([{ sheet: 'S', tabColor: '#00B050' }])
    expect((await s.getBook()).sheets[0]).toMatchObject({ tabColor: '#00B050', view: { gridlines: false } })
    await s.setLayout([{ sheet: 'S', tabColor: null }])
    expect('tabColor' in (await s.getBook()).sheets[0]).toBe(false)
    await s.undo()
    expect((await s.getBook()).sheets[0]).toMatchObject({ tabColor: '#00B050' })
    await expect(s.setLayout([{ sheet: 'S', tabColor: 5 as never }])).rejects.toThrow(InputError)
    await expect(s.setLayout([{ sheet: 'S', view: null } as never])).rejects.toThrow(InputError)
  })
})

describe('files', () => {
  it('writeXlsx writes them, readXlsx reads them, writeKept changes them in place', async () => {
    const data = (await writeXlsx({ book: book() })).data
    expect((await readXlsx(data)).book.sheets[0]).toMatchObject({ tabColor: '#FF0000', view: { gridlines: false } })
    const read = await openKept(data)
    const kept = await writeKept(data, { book: { ...read.book, sheets: [{ ...read.book.sheets[0], tabColor: '#0070C0', view: { headings: false } }] } })
    expect((await readXlsx(kept.data)).book.sheets[0]).toMatchObject({ tabColor: '#0070C0', view: { headings: false } })
  })
})

describe("the book's calculation option", () => {
  it('is checked, left out of the hash, written, read and kept over the file', async () => {
    const b: BookInput = { calculation: 'manual', sheets: [{ name: 'S', cells: { A1: { formula: '1+1' } } }] }
    expect(compute(b).hash).toBe(compute({ sheets: b.sheets }).hash)
    expect(() => compute({ ...b, calculation: 'auto' as never })).toThrow(new InputError('book.calculation', 'expected one of "manual", "autoNoTable", got the string "auto"'))
    const data = (await writeXlsx({ book: b })).data
    expect((await readXlsx(data)).book.calculation).toBe('manual')
    const read = await openKept(data)
    const { calculation: _, ...auto } = read.book
    expect((await readXlsx((await writeKept(data, { book: { ...auto, calculation: 'autoNoTable' } })).data)).book.calculation).toBe('autoNoTable')
    expect((await readXlsx((await writeKept(data, { book: auto })).data)).book.calculation).toBeUndefined()
  })
})

describe('a note shown all the time', () => {
  it('is checked, written, read and kept over the file where it alone changed', async () => {
    expect(() => compute({ sheets: [{ name: 'S', cells: { A1: { note: { text: 'x', shown: 1 as never } } } }] })).toThrow(new InputError('book.sheets[0].cells["A1"].note.shown', 'expected a boolean, got number 1'))
    const b: BookInput = { sheets: [{ name: 'S', cells: { A1: { note: { text: 'x', shown: true } }, B1: { note: { text: 'y' } } } }] }
    const data = (await writeXlsx({ book: b })).data
    expect((await readXlsx(data)).book.sheets[0].cells).toEqual(b.sheets[0].cells)
    const read = await openKept(data)
    const cells = { A1: { note: { text: 'x' } }, B1: { note: { text: 'y', shown: true } } }
    const kept = await writeKept(data, { book: { ...read.book, sheets: [{ ...read.book.sheets[0], cells }] } })
    expect((await readXlsx(kept.data)).book.sheets[0].cells).toEqual(cells)
  })
})
