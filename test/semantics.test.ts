/**
 * Семантика белого списка: ответ по документации Excel; оракул LibreOffice (если есть `soffice`)
 * подтверждает его, кроме записанных расхождений самого LibreOffice.
 */
import { beforeAll, describe, expect, it } from 'vitest'
import { compute, type CellValue } from '../src/index'
import { hasSoffice, oracleValues } from './oracle'
import { CASES, semanticsBook } from './semantics-cases'

/** Где неправ сам LibreOffice (ответ взят по документации Excel). */
const ORACLE_WRONG: Record<string, string> = {
  d5: 'LibreOffice не повторяет ложное 29.02.1900: до 01.03.1900 серийный номер на 1 больше',
  u1: 'в LibreOffice логическое это число: TRUE в диапазоне считается',
  u2: 'LibreOffice по умолчанию не приводит текст аргумента к числу (#VALUE!)',
}

const same = (a: unknown, b: unknown) => (typeof a === 'number' && typeof b === 'number' ? Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a)) : JSON.stringify(a) === JSON.stringify(b))
const book = semanticsBook(CASES)
const clock = { now: '2026-10-04T12:00:00Z', seed: 1 }

describe('семантика белого списка', () => {
  const cells = compute(book, { clock }).sheets.find((s) => s.name === 'Cases')!.cells
  for (const [i, c] of CASES.entries()) {
    it(`${c.id} ${c.formula}${c.note ? ` (${c.note})` : ''}`, () => {
      expect(cells[`A${i + 1}`].value).toEqual(c.excel)
    })
  }

  describe.skipIf(!hasSoffice())('оракул LibreOffice', () => {
    let oracle: Map<string, CellValue>
    beforeAll(async () => {
      oracle = (await oracleValues(book, 'Cases'))!
    }, 120_000)
    it('подтверждает ответы, кроме своих записанных расхождений', () => {
      const off = CASES.filter((c, i) => !same(oracle.get(`A${i + 1}`) ?? null, c.excel)).map((c) => c.id)
      expect(off.sort()).toEqual(Object.keys(ORACLE_WRONG).sort())
    })
  })
})
