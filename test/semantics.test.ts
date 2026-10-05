/**
 * Whitelist semantics: the answer per Excel documentation; the LibreOffice oracle (when `soffice` is
 * available) confirms it, except for the recorded cases where LibreOffice itself differs from Excel.
 */
import { beforeAll, describe, expect, it } from 'vitest'
import { compute, type CellValue } from '../src/index.js'
import { hasSoffice, oracleValues } from './oracle.js'
import { CASES, semanticsBook } from './semantics-cases.js'

/** Where LibreOffice itself differs from Excel (the answer follows Excel documentation). */
const ORACLE_WRONG: Record<string, string> = {
  d5: 'LibreOffice does not reproduce the phantom 1900-02-29: before 1900-03-01 its serial numbers are 1 higher',
  u1: 'in LibreOffice a boolean is a number: TRUE in a range is counted',
  u2: 'by default LibreOffice does not coerce a text argument to a number (#VALUE!)',
}

const same = (a: unknown, b: unknown) => (typeof a === 'number' && typeof b === 'number' ? Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a)) : JSON.stringify(a) === JSON.stringify(b))
const book = semanticsBook(CASES)
const clock = { now: '2026-10-04T12:00:00Z', seed: 1 }

describe('whitelist semantics', () => {
  const cells = compute(book, { clock }).sheets.find((s) => s.name === 'Cases')!.cells
  for (const [i, c] of CASES.entries()) {
    it(`${c.id} ${c.formula}${c.note ? ` (${c.note})` : ''}`, () => {
      expect(cells[`A${i + 1}`].value).toEqual(c.excel)
    })
  }

  describe.skipIf(!hasSoffice())('LibreOffice oracle', () => {
    let oracle: Map<string, CellValue>
    beforeAll(async () => {
      oracle = (await oracleValues(book, 'Cases'))!
    }, 120_000)
    it('confirms the answers, except for its recorded differences', () => {
      const off = CASES.filter((c, i) => !same(oracle.get(`A${i + 1}`) ?? null, c.excel)).map((c) => c.id)
      expect(off.sort()).toEqual(Object.keys(ORACLE_WRONG).sort())
    })
  })
})
