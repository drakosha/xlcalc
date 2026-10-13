/** A selection's copy text is prepared in the library, against stored Excel copies, in all three session kinds. */
import { readFileSync } from 'node:fs'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { clipboardCases, quoteCases } from '../bench/coverage/clipboard/cases.js'
import { CalcError, createInlineCalc, createSession, InputError, type BookInput } from '../src/index.js'
import { CLOCK } from './async-books.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'
import './locales.js'

afterAll(closeCalcs)

const stored = (path: string) => JSON.parse(readFileSync(new URL(path, import.meta.url), 'utf8')) as { answers: Record<string, { text: string }> }

describe.each(['closures', 'eval'] as const)('copy text, %s', (execution) => {
  describe.each(KINDS)('%s', (kind) => {
    let open: Awaited<ReturnType<typeof opener>>
    beforeAll(async () => { open = await opener(kind) })
    for (const tag of ['en-US', 'ru-RU']) for (const platform of ['Mac', 'Windows']) {
      const path = platform === 'Mac' ? `./excel/clipboard/copy-${tag}.json` : `./excel/win/clipboard/${tag === 'ru-RU' ? 'ru-RU/' : ''}copy-held.json`
      const answers = stored(path).answers
      for (const c of clipboardCases(tag)) it(`${platform}, ${tag}: ${c.id}`, async () => {
        // The Mac measurements used dot decimals in both interface languages; the Windows ones used the locale's region.
        const book: BookInput = { ...c.book, locale: platform === 'Mac' ? 'en-US' : tag }
        const sheet = book.sheets[0]
        if (c.id === 'hidden-row') sheet.rows = { 2: { hidden: true } }
        if (c.id === 'hidden-column') sheet.columns = { B: { hidden: true } }
        if (c.id === 'filtered') { sheet.filtered = true; sheet.rows = { 3: { hidden: true } } }
        const s = await open(book, { clock: CLOCK, execution })
        expect(await s.copyText('S', c.range, { trailingNewline: platform === 'Windows' })).toBe(answers[c.id].text)
        await s.close()
      })
    }

    it('copies values by inherited and conditional formats, even without the display, and does not edit the session', async () => {
      const s = await open({ locale: 'ru-RU', sheets: [{ name: 'S', columns: { A: { width: 1, numberFormat: '0.0%' }, B: { width: 1 } }, rows: { 2: { numberFormat: '0.00' } }, cells: { A1: { value: 0.125 }, A2: { value: 0.5 }, B1: { value: true }, B2: { formula: '1/0' }, A3: { value: 2, numberFormat: '0' } }, conditionalFormats: [{ type: 'cellIs', ranges: ['A3'], operator: 'greaterThan', formulas: ['0'], style: { numberFormat: '0.000' } }] }] }, { clock: CLOCK, execution, display: false })
      const before = await s.getSnapshot()
      expect(await s.copyText('s', 'A1:B3', { trailingNewline: false })).toBe('12,5%\tИСТИНА\r\n0,50\t#ДЕЛ/0!\r\n2,000\t')
      const after = await s.getSnapshot()
      expect(after.book).toEqual(before.book)
      expect({ ...after.result, ms: 0 }).toEqual({ ...before.result, ms: 0 })
      expect(s.revision).toBe(0)
      expect(s.history.undo).toBe(0)
      await s.close()
    })

    it('keeps selected blank cells, trims whole rows and columns, and skips zero-sized lines', async () => {
      const s = await open({ sheets: [{ name: 'S', cells: { A1: { value: 'a' }, B1: { value: 'b' }, A3: { value: 'c' } }, rows: { 2: { height: 0 } }, columns: { B: { width: 0 } } }] }, { clock: CLOCK, execution })
      expect(await s.copyText('S', 'A1:C3', { trailingNewline: false })).toBe('a\t\r\nc\t')
      expect(await s.copyText('S', 'A:A', { trailingNewline: false })).toBe('a\r\nc')
      expect(await s.copyText('S', '1:1', { trailingNewline: false })).toBe('a')
      expect(await s.copyText('S', 'B1:B3')).toBe('')
      await s.close()
    })

    it('copies quoted fields and trailing blank cells as the recorded plain-text copies', async () => {
      const answers = stored('./excel/win/clipboard/ru-RU/quotes-held.json').answers as unknown as Record<string, [string, unknown]>
      for (const c of quoteCases()) {
        const s = await open(c.book, { clock: CLOCK, execution })
        expect(await s.copyText('S', c.range)).toBe(answers[c.id][0])
        await s.close()
      }
    })

    it('uses the book date system and copies formula results after edits and undo', async () => {
      const s = await open({ dateSystem: 1904, sheets: [{ name: 'S', columns: { A: { width: 1 } }, cells: { A1: { value: 0, numberFormat: 'yyyy-mm-dd' }, A2: { value: 2 }, A3: { formula: 'A2*2' } } }] }, { clock: CLOCK, execution })
      expect(await s.copyText('S', 'A1')).toBe('1904-01-01\r\n')
      await s.setCells([{ sheet: 'S', cell: 'A2', input: { value: 3 } }])
      expect(await s.copyText('S', 'A3')).toBe('6\r\n')
      await s.undo()
      expect(await s.copyText('S', 'A3')).toBe('4\r\n')
      await s.close()
    })
  })
})

it('refuses an oversized copy instead of silently truncating it, and validates all inputs', () => {
  const s = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 'a' } } }] })
  expect(() => s.copyText('S', 'A1:B2', { maxCells: 3 })).toThrow(CalcError)
  expect(s.copyText('S', 'A1:B2', { maxCells: 4 })).toBe('a\t\r\n\t\r\n')
  expect(() => s.copyText('S', 'A1:XFC2')).toThrow(CalcError)
  expect(s.copyText('S', 'A1:XFD1048576')).toBe('a\r\n')
  expect(() => s.copyText('missing', 'A1')).toThrow(InputError)
  expect(() => s.copyText('S', 'A0')).toThrow(InputError)
  for (const opts of [null, { maxCells: 0 }, { maxCells: 1.5 }, { maxChars: 0 }, { maxChars: 1.5 }, { trailingNewline: 1 }, { unknown: true }]) expect(() => s.copyText('S', 'A1', opts as never)).toThrow(InputError)
  expect(s.revision).toBe(0)
  s.close()
  expect(() => s.copyText('S', 'A1')).toThrow('session is closed')
})

it('bounds encoded text, counting quotes, tabs and CRLF, before joining the rows', () => {
  const s = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 'a\t"b' } } }] })
  const text = '"a\t""b"\t\r\n\t\r\n'
  expect(s.copyText('S', 'A1:B2', { maxChars: text.length })).toBe(text)
  expect(() => s.copyText('S', 'A1:B2', { maxChars: text.length - 1 })).toThrow(CalcError)
  expect(() => s.copyText('S', 'A1', { maxChars: 2 })).toThrow(CalcError)
})

it('accepts cancellation on the asynchronous read without changing the book', async () => {
  const calc = createInlineCalc({ sliceMs: 1 })
  try {
    const s = await calc.createSession({ sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] })
    const controller = new AbortController()
    controller.abort()
    await expect(s.copyText('S', 'A1', { signal: controller.signal })).rejects.toMatchObject({ code: 'aborted' })
    expect(await s.copyText('S', 'A1')).toBe('1\r\n')
    expect(s.revision).toBe(0)
  } finally { await calc.close() }
})
