/** Reproductions from the dependency, session, table and xlsx review, without driving Excel. */
import { describe, expect, it, vi } from 'vitest'
import { compute, createSession, editStructure, validateFormula, type BookInput } from '../src/index.js'
import { writeXlsx } from '../src/xlsx.js'
import { applyEdits, createState } from '../src/engine.js'
import { run } from '../src/exec/budget.js'
import { buildResult } from '../src/result.js'
import { inputScalar } from '../src/internal.js'
import { read, sheet } from './xlsx-fixtures.js'
import './locales.js'

const book = (cells: BookInput['sheets'][number]['cells']): BookInput => ({ sheets: [{ name: 'S', cells }] })
const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

it('keeps inputScalar usable as an array callback without treating its index as a date epoch', () => {
  expect([{ date: '1904-01-01' }, { date: '1904-01-01' }].map(inputScalar)).toEqual([1462, 1462])
})

it('recognizes the Firefox stack-overflow error without masking CSP errors', () => {
  const b = book({ A1: { formula: '1+1' } })
  try {
    vi.stubGlobal('Function', function () { throw Object.assign(new Error('too much recursion'), { name: 'InternalError' }) })
    expect(compute(b, { execution: 'eval' }).sheets[0].cells.A1.value).toBe(2)
    vi.stubGlobal('Function', function () { throw new EvalError('blocked') })
    expect(() => compute(b, { execution: 'eval' })).toThrow(EvalError)
  } finally {
    vi.unstubAllGlobals()
  }
})

it('keeps the pre-optimization hash for rounded values, address order, Unicode and spill fills', () => {
  const b: BookInput = { dateSystem: 1904, locale: 'ru-RU', sheets: [
    { name: 'Я"', cells: { Z12: { value: 999999999999999.6 }, A11: { value: 1.2345678901234567e-20 }, A2: { value: -0 }, A1: { value: 'текст\n"' }, B2: { value: true }, C3: { formula: '1/0' }, B1: { value: 0.1 + 0.2 }, D9: { value: null } } },
    { name: 'Fill', cells: { A1: { formula: 'SEQUENCE(100000,1,0,0)' } } },
  ] }
  // Recorded with commit 8ad9e3b before the result/canonicalization changes.
  expect(compute(b, { clock }).hash).toBe('3511b1dd8660658408ec209be8b9266fdb25bb36c467e40c9abe469e8b3886d0')
})

it('a settled spill is not recomputed on an unrelated edit, but blockers and released cells are observed', () => {
  const b = book({ A1: { formula: 'SEQUENCE(2)' }, C1: { formula: 'SUM(A1#)' }, Z1: { value: 0 } })
  const state = createState(b, { now: 46300.5, seed: clock.seed })
  const before = state.spills.step
  run(applyEdits(state, [{ sheet: 'S', cell: 'Z1', input: { value: 1 } }]))
  expect(state.spills.step).toBe(before)
  for (const input of [{ value: 9 }, null]) {
    run(applyEdits(state, [{ sheet: 'S', cell: 'A2', input }]))
    const r = run(buildResult(state, clock, 0, false))
    expect(r.sheets[0].cells.A1.value).toEqual(input ? { error: 'SPILL' } : 1)
    expect(r.hash).toBe(compute(state.input, { clock, display: false }).hash)
  }
})

describe.each(['closures', 'eval'] as const)('review regressions, %s', (execution) => {
  it('changing a serial into a typed date uses the book epoch for content equality', () => {
    const s = createSession({ ...book({ A1: { value: 1462 }, B1: { formula: 'A1*2' } }), dateSystem: 1904 }, { execution, clock })
    for (const input of [{ value: { date: '1904-01-01' } }, { value: 1462 }]) {
      s.setCells([{ sheet: 'S', cell: 'A1', input }])
      const full = compute(s.getBook(), { execution, clock })
      expect(s.getResult().hash).toBe(full.hash)
      expect(s.getResult().sheets).toEqual(full.sheets)
    }
    s.close()
  })

  it('changing only the fixed array area releases and restores spill cells', () => {
    const s = createSession(book({ A1: { formula: 'SEQUENCE(2)' }, C1: { formula: 'SUM(A1:A2)' } }), { execution, clock })
    for (const input of [{ formula: 'SEQUENCE(2)', array: 'A1' }, { formula: 'SEQUENCE(2)' }]) {
      s.setCells([{ sheet: 'S', cell: 'A1', input }])
      const full = compute(s.getBook(), { execution, clock })
      expect(s.getResult().hash).toBe(full.hash)
      const fresh = createSession(s.getBook(), { execution, clock })
      expect(s.getCells('S', 'A1:C2')).toEqual(fresh.getCells('S', 'A1:C2'))
      fresh.close()
    }
    s.close()
  })

  it('a range through unary plus orders formulas in its interior and listens to their edits', () => {
    const b = book({ A1: { value: 1 }, A2: { value: 2 }, B1: { formula: '10' }, C1: { value: 3 }, D1: { formula: 'SUM((+A1:A2):C1)' } })
    expect(compute(b, { execution }).sheets[0].cells.D1.value).toBe(16)
    const s = createSession(b, { execution, clock })
    for (const input of [{ value: 20 }, { formula: '30' }]) {
      s.setCells([{ sheet: 'S', cell: 'B1', input }])
      expect(s.getResult().hash).toBe(compute(s.getBook(), { execution, clock }).hash)
    }
    s.close()
  })

  it.each(['SUM(A1:LET(x,C1,x))', 'SUM(LET(x,A1,x):C1)', 'LET(z,A1,SUM(z:LET(x,C1,x)))', 'SUM(A1:LET(x,B1,y,C1,y))'])('LET range ends keep their local bindings: %s', (formula) => {
    const b = book({ A1: { value: 1 }, B1: { formula: '10' }, C1: { value: 3 }, D1: { formula } })
    expect(compute(b, { execution }).sheets[0].cells.D1.value).toBe(14)
  })

  it('both spilled ends contribute to the whole range, including the diagonal corner', () => {
    const b = book({ A1: { formula: 'SEQUENCE(5)' }, C1: { formula: 'SEQUENCE(1,5)' }, D4: { formula: 'I1+1' }, I1: { formula: 'SUM(A1#:C1#)' } })
    const r = compute(b, { execution })
    expect(r.sheets[0].cells.I1.value).toEqual({ error: 'CYCLE' })
    expect(r.sheets[0].cells.D4.value).toEqual({ error: 'CYCLE' })
  })

  it('a long flat formula keeps the evaluation order and the execution modes agree', () => {
    const sum = Array(2000).fill('A1').join('+')
    expect(validateFormula(sum).ok).toBe(true)
    const b = book({ A1: { value: 1 }, B1: { formula: sum }, C1: { formula: `IF(FALSE,${sum},42)` }, D1: { formula: `LET(x,2,${Array(2000).fill('x').join('+')})` } })
    const r = compute(b, { execution })
    expect([r.sheets[0].cells.B1.value, r.sheets[0].cells.C1.value, r.sheets[0].cells.D1.value]).toEqual([2000, 42, 4000])
    const s = createSession(book({ A1: { value: 2 } }), { execution, clock })
    s.setCells([{ sheet: 'S', cell: 'B1', input: { formula: sum } }])
    expect(s.getResult().sheets[0].cells.B1.value).toBe(4000)
    expect(s.getResult().hash).toBe(compute(s.getBook(), { execution, clock }).hash)
    s.close()
  })
})

it('calculated column formulas follow moved tables and are copied relative to new rows', () => {
  const b = book({ A1: { value: 'Qty' }, B1: { value: 'Double' }, A2: { value: 3 }, B2: { formula: 'A2*2+$A$2' }, A3: { value: 7 }, B3: { formula: 'A3*2+$A$2' }, A4: { value: 11 }, B4: { formula: 'A4*2+$A$2' } })
  b.sheets[0].tables = [{ name: 'Sales', range: 'A1:B4', columns: [{ name: 'Qty' }, { name: 'Double', formula: 'A2*2+$A$2' }] }]
  const inside = editStructure(b, { op: 'insertRows', sheet: 'S', row: 3, count: 2 })
  if (!inside.ok) throw new Error(inside.message)
  expect(inside.book.sheets[0].cells.B3.formula).toBe('A3*2+$A$2')
  expect(inside.book.sheets[0].cells.B4.formula).toBe('A4*2+$A$2')
  const moved = editStructure(b, { op: 'insertRows', sheet: 'S', row: 1, count: 1 })
  if (!moved.ok) throw new Error(moved.message)
  expect(moved.book.sheets[0].tables![0].columns[1].formula).toBe('A3*2+$A$3')
  const grown = editStructure(moved.book, { op: 'insertRows', sheet: 'S', row: 4, count: 1 })
  if (!grown.ok) throw new Error(grown.message)
  expect(grown.book.sheets[0].cells.B4.formula).toBe('A4*2+$A$3')
  const first = editStructure(b, { op: 'insertRows', sheet: 'S', row: 2, count: 2 })
  if (!first.ok) throw new Error(first.message)
  expect(first.book.sheets[0].tables![0].columns[1].formula).toBe('A2*2+$A$4')
  expect(first.book.sheets[0].cells.B3.formula).toBe('A3*2+$A$4')
  const deleted = editStructure(b, { op: 'deleteRows', sheet: 'S', row: 2, count: 1 })
  if (!deleted.ok) throw new Error(deleted.message)
  expect(deleted.book.sheets[0].tables![0].columns[1].formula).toBe('A2*2+#REF!')
})

it('reports a calculated column formula that cannot be written to xlsx', async () => {
  const b = book({ A1: { value: 'X' }, A2: { value: 1 } })
  b.sheets[0].tables = [{ name: 'Sales', range: 'A1:A2', columns: [{ name: 'X', formula: 'Missing!A2' }] }]
  const r = await writeXlsx(b)
  expect(r.issues).toContainEqual(expect.objectContaining({ code: 'formula', sheet: 'S' }))
})

it.each(['row', 'column'] as const)('reports properties lost from a %s style', async (kind) => {
  const data = '<c r="A1"><v>42</v></c>'
  const body = kind === 'row' ? `<sheetData><row r="1" s="1" customFormat="1">${data}</row></sheetData>` : `<cols><col min="1" max="1" style="1" width="12" customWidth="1"/></cols><sheetData><row r="1">${data}</row></sheetData>`
  const r = await read({ sheets: [{ name: 'S', body }], styles: '<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="0" applyAlignment="1"><alignment textRotation="45"/></xf></cellXfs>' })
  expect(r.issues).toContainEqual(expect.objectContaining({ code: 'style', sheet: 'S', cells: [kind === 'row' ? '1:1' : 'A:A'] }))
})

it('reads a shared formula group larger than the JS argument limit', async () => {
  const n = 150_000
  const rows = Array.from({ length: n }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><f t="shared" si="0"${i === 0 ? ` ref="A1:A${n}">ROW()</f>` : '/>'}<v>${i + 1}</v></c></row>`).join('')
  const r = await sheet(rows)
  expect(Object.keys(r.book.sheets[0].cells)).toHaveLength(n)
  expect(r.book.sheets[0].cells[`A${n}`].formula).toBe('ROW()')
}, 60_000)
