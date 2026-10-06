/**
 * The two execution modes against each other (`execution-compare.ts`) on every case book: the coverage
 * measurement, the case sets of `docs/excel-cases/`, the probes, the measurement and the sets in ru-RU, the
 * semantics cases, the synthetic book, the seeded books of the structure tests, the spill cases and the spill
 * book with its seeded edits, and a book of every kind of compiled node. Values (cell by cell, without
 * tolerance: `Object.is`), issues with their reasons and check codes, volatile cells and hashes are equal.
 */
import { runInContext } from 'node:vm'
import { describe, expect, it } from 'vitest'
import { buildProbes, PROBE_SHEETS } from '../bench/coverage/probes.js'
import type { BookInput, CellInput } from '../src/index.js'
import { synthBook } from '../src/synth.js'
import { CASES, excelBook } from './excel-cases.js'
import { RU_SETS } from './excel-ru-cases.js'
import { caseBook, loadSet, SETS } from './excel-sets.js'
import { byClosures, byEval, ctx, expectSame, strict } from './execution-compare.js'
import { randomBook } from './seeded-books.js'
import { CASES as SEMANTICS, semanticsBook } from './semantics-cases.js'
import { apply } from './session-edits.js'
import { spillBook, spillEdits } from './session-spill-edits.js'
import type { SpillCase } from '../bench/coverage/spill/cases.js'

const BOOKS: [string, BookInput][] = [
  ['the coverage measurement', excelBook(CASES)],
  ...(await Promise.all(SETS.map(async (s): Promise<[string, BookInput]> => {
    const mod = await loadSet(s.name)
    return [`the case set ${s.name}`, caseBook(mod.buildProbes(), mod.SHEETS)]
  }))),
  ['the probes', caseBook(buildProbes(), PROBE_SHEETS)],
  ['the coverage measurement in ru-RU', { ...excelBook(CASES), locale: 'ru-RU' }],
  ...(await Promise.all(RU_SETS.map(async (s): Promise<[string, BookInput]> => {
    const mod = await loadSet(s.name)
    return [`the case set ${s.name} in ru-RU`, { ...caseBook(mod.buildProbes(), mod.SHEETS), locale: 'ru-RU' }]
  }))),
  ['the semantics cases', semanticsBook(SEMANTICS)],
  ['the synthetic book', synthBook({ seed: 1, rows: 3400 }).book],
  ['the spill book', spillBook()],
]
const SPILL_CASES = (await Promise.all(['base', 'order', 'held'].map(async (set) => ((await import(`../bench/coverage/spill/${set}.ts`)) as { buildSpillCases: () => SpillCase[] }).buildSpillCases()))).flat()

/** Every kind of node of the compiler, lifted and not; names, an external function, a cycle, SUBTOTAL over SUBTOTAL. */
const f = (formula: string): CellInput => ({ formula })
/** B2:B7: each comparison below, at and above A1, scalar and lifted. */
const COMPARISONS = Object.fromEntries(['=', '<>', '<', '>', '<=', '>='].map((op, i) => [`B${i + 2}`, f(`(A1${op}2)&(A1${op}3)&(A1${op}4)&CONCAT(A1:A4${op}3)`)]))
const NODES: BookInput = {
  sheets: [
    {
      name: 'N',
      cells: {
        A1: { value: 3 }, A2: { value: '5' }, A3: { value: true }, A4: { value: -2.5 }, A5: f('SUBTOTAL(9,A1:A4)'),
        B1: f('1+2-3*4/5^2&"x"'),
        ...COMPARISONS,
        B8: f('-A1%'), B9: f('SUM(-A1:A4)+SUM(A1:A4%)'), B10: f('SUM(A1:A4=3)+SUM((A1:A4>0)*A1:A4)'), B11: f('SUM(A1:A4&"")'),
        B12: f('#N/A'), B13: f('IFERROR(SUM({1,#N/A}),"e")'), B14: f('SUM({1,2;3,4}*{10;100})'), B15: f('A1:A1*2+SUM(A1)+SUM(+A1:A4)+SUM((A1:A4))'),
        B16: f('ROUND(A4,0)+SUM(ROUND(A1:A4/3,1))'), B17: f('IF(A1>0,"yes",1/0)'), B18: f('SUM(IF(A1:A4>1,A1:A4,0))'), B19: f('IFERROR(1/0,IF(A3,2))'),
        B20: f('RAND()+RANDBETWEEN(1,6)+SUM(RANDARRAY(2,2))'), B21: f('TODAY()+NOW()'), B22: f('rate*2+SUM(span)+local'), B23: f('BALANCE("2025-01")'),
        B24: f('BALANCE("2025-02")'), B25: f('A1:A4*2'), B26: f('NOPE(1)'), B27: f('SUBTOTAL(9,A1:A5)+AGGREGATE(9,0,A1:A5)'), B28: f('SUM(B:B*0)'),
        B29: f('B12+1'), B30: f('INDEX(A1:A4,2)&COUNTA(UNIQUE(A1:A4))'), C1: f('C2+1'), C2: f('C1+1'), C3: f('C1*2'),
        // spills: a spill reference read, lifted and in a name
        D1: f('SEQUENCE(2,2)'), D4: f('D1#*2'), D7: f('SUM(D1#)+ROWS(sq)+SUM(SEQUENCE(1,2)*D4#)'),
      },
    },
    { name: 'M', cells: { A1: { value: 0.07 }, B1: f('local*10') } },
  ],
  names: [
    { name: 'rate', ref: "'N'!$A$1" },
    { name: 'span', ref: "'N'!$A$1:$A$4" },
    { name: 'local', ref: "'N'!$A$4", scope: 'N' },
    { name: 'local', ref: "'M'!$A$1", scope: 'M' },
    { name: 'sq', ref: "'N'!$D$1#" },
  ],
}
const EXTERNALS = { functions: ['BALANCE'], values: { 'BALANCE("2025-01")': 1250 } }

describe('execution modes against each other', () => {
  it('the closures mode computes where code generation from strings is disabled, and the eval mode cannot', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { formula: '1+1' } } }] }
    expect(() => runInContext('new Function("return 1")', ctx)).toThrow()
    expect(() => strict.compute(book, { execution: 'eval' })).toThrow(/new Function/)
    expect(strict.compute(book, { execution: 'closures' }).sheets[0].cells.A1.value).toBe(2)
  })

  for (const [name, book] of BOOKS) it(`${name}: the same values, issues and hash`, () => expectSame(byEval(book), byClosures(book)), 60_000)

  it(`the ${SPILL_CASES.length} books of the spill cases`, () => {
    for (const c of SPILL_CASES) expectSame(byEval(c.book), byClosures(c.book))
  })

  it('200 seeded edits of the spill book: the books after each one', () => {
    let book = spillBook()
    for (const batch of spillEdits(11, 200)) {
      book = apply(book, batch)
      expectSame(byEval(book), byClosures(book))
    }
  }, 120_000)

  it('40 seeded books of the structure tests', () => {
    for (let seed = 1; seed <= 40; seed++) expectSame(byEval(randomBook(seed)), byClosures(randomBook(seed)))
  })

  it('every kind of compiled node', () => {
    const r = byEval(NODES, { externals: EXTERNALS })
    expectSame(r, byClosures(NODES, { externals: EXTERNALS }))
    // the book reaches what it is for: a cycle, an array result, a missing external value, an invalid formula
    expect(r.issues.map((i) => `${i.cell} ${i.error}`)).toEqual(expect.arrayContaining(['B24 NA', 'B25 SPILL', 'B26 NAME', 'C1 CYCLE']))
    expect([r.sheets[0].cells.D4.spill, r.sheets[0].cells.D7.value]).toEqual(['D4:E5', 10 + 2 + (2 + 8 + 6 + 16)])
    expect(r.issues.find((i) => i.cell === 'B24')?.reason).toBe('value not found (or no value for an external call)')
  })
})
