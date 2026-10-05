/**
 * Whitelist semantics: the answer per Excel documentation for every case, and the LibreOffice oracle
 * confirms it, except for the recorded cases where LibreOffice itself differs from Excel. The oracle's
 * answers are stored in `test/oracle/libreoffice.json`, so this test needs no office suite;
 * `npm run test:oracle` recalculates them with `soffice` and rewrites the file.
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { beforeAll, describe, expect, it } from 'vitest'
import { compute, type CellValue } from '../src/index.js'
import { inputHash } from '../src/canon.js'
import { CASES, semanticsBook } from './semantics-cases.js'

/** Where LibreOffice itself differs from Excel (the answer follows Excel documentation). */
const ORACLE_WRONG: Record<string, string> = {
  d5: 'LibreOffice does not reproduce the phantom 1900-02-29: before 1900-03-01 its serial numbers are 1 higher',
  u1: 'in LibreOffice a boolean is a number: TRUE in a range is counted',
  u2: 'by default LibreOffice does not coerce a text argument to a number (#VALUE!)',
}

/** The stored oracle answers: which LibreOffice, the input hash of the semantics book, answers by case. */
interface OracleFile {
  generator: string
  input: string
  cases: Record<string, { formula: string; value: CellValue }>
}

const ORACLE_FILE = new URL('./oracle/libreoffice.json', import.meta.url)
const UPDATE = process.env.FORMULA_ORACLE === 'update'
const same = (a: unknown, b: unknown) => (typeof a === 'number' && typeof b === 'number' ? Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a)) : JSON.stringify(a) === JSON.stringify(b))
const book = semanticsBook(CASES)
const clock = { now: '2026-10-04T12:00:00Z', seed: 1 }

/** Recalculates the semantics book with soffice and rewrites the stored answers. */
async function recalculate(): Promise<OracleFile> {
  const { oracleValues, sofficeVersion } = await import('./oracle/libreoffice.js')
  const generator = sofficeVersion()
  if (!generator) throw new Error('npm run test:oracle needs soffice (LibreOffice) on PATH')
  const values = await oracleValues(book, 'Cases')
  const file: OracleFile = {
    generator: `${generator}, soffice --headless, en-US profile`,
    input: inputHash(book),
    cases: Object.fromEntries(CASES.map((c, i) => [c.id, { formula: c.formula, value: values.get(`A${i + 1}`) ?? null }])),
  }
  writeFileSync(ORACLE_FILE, JSON.stringify(file, null, 2) + '\n')
  return file
}

describe('whitelist semantics', () => {
  const cells = compute(book, { clock }).sheets.find((s) => s.name === 'Cases')!.cells
  for (const [i, c] of CASES.entries()) {
    it(`${c.id} ${c.formula}${c.note ? ` (${c.note})` : ''}`, () => {
      expect(cells[`A${i + 1}`].value).toEqual(c.excel)
    })
  }

  describe('LibreOffice oracle', () => {
    let oracle: OracleFile
    beforeAll(async () => {
      oracle = UPDATE ? await recalculate() : (JSON.parse(readFileSync(ORACLE_FILE, 'utf8')) as OracleFile)
    }, 120_000)

    it('the stored answers are for the current cases (otherwise run npm run test:oracle)', () => {
      expect(Object.keys(oracle.cases).sort()).toEqual(CASES.map((c) => c.id).sort())
      for (const c of CASES) expect(oracle.cases[c.id].formula, c.id).toBe(c.formula)
      expect(oracle.input, 'the semantics book changed').toBe(inputHash(book))
    })

    it('confirms the answers, except for its recorded differences', () => {
      const off = CASES.filter((c) => !same(oracle.cases[c.id]?.value ?? null, c.excel)).map((c) => c.id)
      expect(off.sort()).toEqual(Object.keys(ORACLE_WRONG).sort())
    })
  })
})
