/**
 * A formula of very many operators: `1+1+1` nests without a parenthesis, so the depth limit does not see it, and every walk
 * over the tree recurses. One such formula must not take down the workbook it is in, wherever it gets in from.
 */
import { describe, expect, it } from 'vitest'
import { compute, createSession } from '../src/index.js'
import { fromLegacy, readXlsx, writeXlsx } from '../src/xlsx.js'
import { sheet } from './xlsx-fixtures.js'

const CLOCK = { now: '2026-10-04T12:00:00Z', seed: 7 }
const terms = (n: number) => Array(n).fill('1').join('+')
const FORMS: [string, (n: number) => string][] = [
  ['+', terms],
  ['%', (n) => `1${'%'.repeat(n - 1)}`],
]
const book = (formula: string) => ({ sheets: [{ name: 'S', cells: { A1: { formula }, B1: { value: 5 }, C1: { formula: 'B1+1' } } }] })

describe('a formula of very many operators', () => {
  it('2,000 operators compute, and every walk over the tree takes them', async () => {
    const formula = terms(2001)
    for (const execution of ['closures', 'eval'] as const) expect(compute(book(formula), { clock: CLOCK, execution }).sheets[0].cells.A1.value).toBe(2001)
    const s = createSession(book(formula), { clock: CLOCK })
    s.editStructure({ op: 'insertRows', sheet: 'S', row: 1 })
    expect(s.getCell('S', 'A2')?.value).toBe(2001)
    expect(fromLegacy(formula)).toBe(formula)
    const { data, issues } = await writeXlsx({ book: book(formula) })
    expect(issues.filter((i) => i.code === 'formula')).toEqual([])
    expect((await readXlsx(data)).book.sheets[0].cells.A1).toEqual({ formula })
  })

  it.each(FORMS)('more, by %s: an error of that cell alone in compute', (_, form) => {
    const r = compute(book(form(2002)), { clock: CLOCK })
    expect(r.sheets[0].cells.A1.value).toEqual({ error: 'ERROR' })
    expect(r.diagnostics[0]).toMatchObject({ cell: 'A1', check: 'too_deep', reason: 'more than 2000 operators in one formula' })
    expect(r.sheets[0].cells.C1.value).toBe(6)
  })

  it.each([2002, 3000, 4096])('%i terms are written as the value, with an issue', async (n) => {
    const result = compute(book(terms(n)), { clock: CLOCK })
    const { data, issues } = await writeXlsx({ book: book(terms(n)), result })
    expect(issues.map((i) => i.code)).toContain('formula')
    expect((await readXlsx(data)).book.sheets[0].cells.B1).toEqual({ value: 5 })
  })

  it('in a file: the file reads, the value the file stores stays, the other cells are whole', async () => {
    const r = await sheet(`<row r="1"><c r="A1"><f>${terms(3000)}</f><v>3000</v></c><c r="B1"><v>5</v></c><c r="C1"><f>B1+1</f><v>6</v></c></row>`)
    expect(r.cached.sheets[0].cells.A1.value).toBe(3000)
    expect(r.book.sheets[0].cells.A1).toEqual({ formula: terms(3000) })
    expect(r.book.sheets[0].cells.B1).toEqual({ value: 5 })
    expect(compute(r.book, { clock: CLOCK }).sheets[0].cells.C1.value).toBe(6)
  })

  it('a host that types one gets a refusal, not a stack overflow', () => {
    const s = createSession(book('1'), { clock: CLOCK })
    const r = s.setCellText('S', 'D1', `=${terms(3000)}`)
    expect(r).toMatchObject({ ok: false, code: 'too_deep' })
  })
})
