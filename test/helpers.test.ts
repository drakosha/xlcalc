/**
 * The helpers that answer by the engine's rules without a book: a sheet name as the book check and `renameSheet` take
 * it, an error as a cell shows it, and values compared and written as the values hash compares and writes them, in both
 * execution modes.
 */
import { describe, expect, it } from 'vitest'
import type { CellValue, ErrorCode, Scalar } from '../src/contract.js'
import { canonValue, compute, displayValue, errorText, InputError, editStructure, sameValue, validateSheetName } from '../src/index.js'
import './locales.js'

describe('validateSheetName', () => {
  it('takes what Excel takes', () => {
    for (const name of ['Sales', 'Q1 data', "O'Brien", 'Лист1', '2024', 'x'.repeat(31), 'Historyx', '#!$%']) expect(validateSheetName(name), name).toEqual({ ok: true })
  })

  it('refuses with a code, a message in the locale and the place of the character', () => {
    expect(validateSheetName('')).toEqual({ ok: false, code: 'empty', message: 'a sheet name must not be empty' })
    expect(validateSheetName('x'.repeat(32))).toMatchObject({ ok: false, code: 'too_long' })
    expect(validateSheetName('Q1/Q2')).toEqual({ ok: false, code: 'character', message: 'a sheet name must not contain "/" (Excel forbids : \\ / ? * [ ])', at: 2 })
    expect(validateSheetName("Sales'")).toMatchObject({ ok: false, code: 'apostrophe', at: 5 })
    expect(validateSheetName("'Sales")).toMatchObject({ ok: false, code: 'apostrophe', at: 0 })
    expect(validateSheetName('HISTORY')).toMatchObject({ ok: false, code: 'reserved' })
    expect(validateSheetName('a*b', { locale: 'ru-RU' })).toEqual({ ok: false, code: 'character', message: 'Имя листа не может содержать «*» (Excel запрещает : \\ / ? * [ ])', at: 1 })
  })

  it('gives the message renameSheet refuses with, and the book check throws', () => {
    const book = { locale: 'ru-RU', sheets: [{ name: 'S', cells: {} }] }
    for (const name of ['', 'x'.repeat(40), 'a[b]', "'x'", 'history']) {
      const r = editStructure(book, { op: 'renameSheet', sheet: 'S', name })
      const v = validateSheetName(name, { locale: 'ru-RU' })
      expect(!r.ok && !v.ok && r.message === v.message, name).toBe(true)
      expect(() => compute({ sheets: [{ name, cells: {} }] }), name).toThrow(InputError)
    }
  })

  it('wrong input is an InputError', () => {
    expect(() => validateSheetName(1 as unknown as string)).toThrow(InputError)
    expect(() => validateSheetName('S', { locale: 'xx' })).toThrow(/unknown locale/)
  })
})

describe('errorText', () => {
  const CODES: ErrorCode[] = ['DIV0', 'NA', 'NAME', 'NUM', 'REF', 'VALUE', 'NULL', 'CYCLE', 'SPILL', 'CALC', 'ERROR']

  it('an error as a cell shows it, in the locale', () => {
    expect(CODES.map((c) => errorText(c))).toEqual(['#DIV/0!', '#N/A', '#NAME?', '#NUM!', '#REF!', '#VALUE!', '#NULL!', '#CYCLE!', '#SPILL!', '#CALC!', '#ERROR!'])
    expect([errorText('DIV0', 'ru-RU'), errorText('NA', 'ru-RU'), errorText('VALUE', 'ru-RU')]).toEqual(['#ДЕЛ/0!', '#Н/Д', '#ЗНАЧ!'])
    for (const locale of ['en-US', 'ru-RU']) for (const c of CODES) expect(errorText(c, locale), `${locale} ${c}`).toBe(displayValue({ error: c }, { locale }).text)
  })

  it('wrong input is an InputError', () => {
    expect(() => errorText('DIV/0' as ErrorCode)).toThrow(/expected an error code/)
    expect(() => errorText('NA', 'xx')).toThrow(InputError)
  })
})

describe('canonValue', () => {
  it('the text of a value in the values hash', () => {
    expect([0.1 + 0.2, -0, 1 / 3, 'a"b', true, null, { error: 'DIV0' } as CellValue].map(canonValue)).toEqual(['n 0.3', 'n 0', 'n 0.333333333333333', 's "a\\"b"', 'b TRUE', '_', 'e DIV0'])
    expect(() => canonValue(Number.NaN)).toThrow(InputError)
  })
})

describe.each(['closures', 'eval'] as const)('sameValue as the values hash compares (%s)', (execution) => {
  const hashOf = (v: CellValue | undefined) => compute({ sheets: [{ name: 'S', cells: v === undefined || v === null ? {} : { A1: { value: v as Scalar } } }] }, { clock: { now: '2026-10-10T12:00:00Z', seed: 1 }, execution }).hash
  const pairs: [CellValue | undefined, CellValue | undefined][] = [
    [0.1 + 0.2, 0.3],
    [1 / 3, 0.333333333333333],
    [-0, 0],
    [123456789012345678, 123456789012346000],
    [1e-16, 0],
    [1, 1.0000000000001],
    [1, true],
    ['a', 'A'],
    ['', null],
    [null, undefined],
    [false, 0],
    ['1', 1],
  ]

  it.each(pairs.map((p) => [JSON.stringify(p), ...p] as const))('%s', (_, a, b) => {
    expect(sameValue(a, b)).toBe(hashOf(a) === hashOf(b))
    expect(sameValue(b, a)).toBe(sameValue(a, b))
  })

  it('errors by their code', () => {
    expect(sameValue({ error: 'NA' }, { error: 'NA' })).toBe(true)
    expect(sameValue({ error: 'NA' }, { error: 'REF' })).toBe(false)
    expect(sameValue({ error: 'NA' }, '#N/A')).toBe(false)
  })
})
