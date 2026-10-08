/**
 * Spilling beyond the Excel cases: `E2#` in the syntax and in other languages, determinism over the input's
 * order, large arrays kept as fills (results, hash, cost).
 */
import { describe, expect, it } from 'vitest'
import { compute, ruRU, toInvariant, toLocal, validateFormula, type BookInput } from '../src/index.js'
import { canonical, parse } from '../src/syntax.js'
import { spillBook } from './session-spill-edits.js'

const clock = { now: '2026-10-07T12:00:00Z', seed: 3 }

/** The book with its sheets' cells in reverse order. */
const reversed = (b: BookInput): BookInput => ({ ...b, sheets: b.sheets.map((s) => ({ ...s, cells: Object.fromEntries(Object.entries(s.cells).reverse()) })) })

describe('spill references in the syntax', () => {
  it('E2# parses, prints back and is valid; # follows a cell only', () => {
    expect(canonical(parse("SUM(e2#)+'Data'!$B$3#*2"))).toBe("SUM(E2#)+'Data'!$B$3#*2")
    expect(validateFormula('S!E2#', { sheets: ['S'] })).toMatchObject({ ok: true, canonical: "'S'!E2#", sheets: ['S'] })
    expect(validateFormula('A:B#')).toMatchObject({ ok: false, code: 'syntax', message: 'a spill reference (#) follows a single cell' })
    // the range operator to a spill reference, as Excel reads it
    expect(validateFormula('A1:B2#')).toMatchObject({ ok: true, canonical: 'A1:B2#' })
  })

  it('a lost spill reference is read and written as Excel writes it', () => {
    for (const f of ['SUM(#REF!#)', "'S'!#REF!#+1", '#REF!E2#', '#REF!$E$2#']) expect(canonical(parse(f))).toBe(f)
    expect(compute({ sheets: [{ name: 'S', cells: { A1: { formula: 'SUM(#REF!#)' } } }] }, { clock }).sheets[0].cells.A1.value).toEqual({ error: 'REF' })
  })

  it('the locale translation passes # after a reference through', () => {
    expect(toLocal("SUM(E2#,'Данные'!B1#)/ROWS(sp)", ruRU)).toEqual({ ok: true, formula: "СУММ(E2#;'Данные'!B1#)/ЧСТРОК(sp)" })
    expect(toInvariant('=СУММ(E2#;1,5)*Лист1!A1#', ruRU)).toEqual({ ok: true, formula: 'SUM(E2#,1.5)*Лист1!A1#' })
  })
})

describe('spills computed', () => {
  it('the same values and hash whatever the order of the input (two arrays over one cell too)', () => {
    const cross: BookInput = { sheets: [{ name: 'S', cells: { C1: { formula: 'SEQUENCE(3)' }, A2: { formula: 'SEQUENCE(1,3)' }, B5: { formula: 'SEQUENCE(4)' }, A7: { formula: 'SEQUENCE(1,4)' } } }] }
    for (const book of [spillBook(), cross]) expect(compute(reversed(book), { clock }).hash).toBe(compute(book, { clock }).hash)
    // of two independent arrays the one further down spills
    expect(compute(cross, { clock }).sheets[0].cells.C1.value).toEqual({ error: 'SPILL' })
  })

  it('a whole column: the data rows are listed, the rest is one fill with one line in the hash', () => {
    const cells: BookInput['sheets'][number]['cells'] = { B1: { formula: 'A:A*2' }, C1: { formula: 'ROWS(B1#)' }, C2: { formula: 'B1048576' }, C3: { formula: 'SUM(B:B)' } }
    for (let r = 1; r <= 100; r++) cells[`A${r}`] = { value: r }
    const t = performance.now()
    const r = compute({ sheets: [{ name: 'S', cells }] }, { clock })
    expect(performance.now() - t).toBeLessThan(1000)
    const out = r.sheets[0].cells
    expect(Object.keys(out).filter((a) => out[a].spillOf === 'B1')).toHaveLength(99)
    expect([out.B1.spill, out.B1.spillFill, out.C1.value, out.C2.value, out.C3.value]).toEqual(['B1:B1048576', 0, 1048576, 0, 10100])
    // the fill is in the hash: another fill value, another hash
    const plus = { ...cells, B1: { formula: 'A:A*2+(ROW(A:A)>100)' } }
    expect(compute({ sheets: [{ name: 'S', cells: plus }] }, { clock }).hash).not.toBe(r.hash)
  })

  it('an array of up to 65,536 values is listed whole, its trailing values too', () => {
    const r = compute({ sheets: [{ name: 'S', cells: { A1: { formula: 'SEQUENCE(65536)*0' } } }] }, { clock })
    expect(Object.keys(r.sheets[0].cells)).toHaveLength(65536)
    expect(r.sheets[0].cells.A1.spillFill).toBeUndefined()
  })
})
