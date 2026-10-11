import { describe, expect, it } from 'vitest'
import { compute, openSession, type BookInput, type CellInput } from '../src/index.js'
import { find } from '../src/functions/lookup/search.js'
import { compareValues, E, type V } from '../src/values.js'
import { formatAddress } from '../src/model.js'
import './locales.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

it('repeated exact numeric lookups do not scan the same immutable line again', () => {
  const values = Array.from({ length: 4096 }, (_, i) => i)
  let comparisons = 0
  const compare = (a: V, b: V) => { comparisons++; return compareValues(a, b) }
  const line = { values, stride: 1, at: (i: number) => values[i], n: values.length, filled: values.length }
  for (const key of [4095, 4094, 4093, 4092, 4091]) expect(find(line, key, 0, false, compare)).toBe(key)
  expect(comparisons).toBeLessThan(values.length * 2)
})

it('a few hits near the start do not read the rest of a long line just to build an index', () => {
  let reads = 0
  const values = new Proxy(Array.from({ length: 4096 }, (_, i) => i), {
    get(target, key, receiver) {
      if (typeof key === 'string' && /^\d+$/.test(key)) reads++
      return Reflect.get(target, key, receiver)
    },
  })
  const line = { values, stride: 1, at: (i: number) => values[i], n: values.length, filled: values.length }
  for (const key of [0, 1, 2, 3]) expect(find(line, key, 0, false, compareValues)).toBe(key)
  expect(reads).toBeLessThan(20)
})

it('an index keeps numeric and boolean types distinct and compares numbers without output rounding', () => {
  const values: V[] = Array.from({ length: 300 }, (_, i) => i + 10)
  values.splice(0, 10, null, E.NA, -0, false, true, '0', 0.1 + 0.2, 0.3, 0, true)
  const line = { values, stride: 1, at: (i: number) => values[i], n: values.length, filled: values.length }
  for (let repeat = 0; repeat < 3; repeat++) for (const [key, pos] of [[0, 2], [false, 3], [true, 4], ['0', 5], [0.1 + 0.2, 6], [0.3, 7]] as const)
    expect(find(line, key, 0, false, compareValues)).toBe(pos)
})

describe.each(['closures', 'eval'] as const)('lookup cache lifetime, %s', (execution) => {
  it('horizontal lookups use the stored width of a whole-row table', () => {
    const cells: Record<string, CellInput> = {}
    for (let c = 0; c < 400; c++) { cells[formatAddress(0, c)] = { value: c + 1 }; cells[formatAddress(1, c)] = { value: (c + 1) * 10 } }
    const s = openSession({ sheets: [{ name: 'Data', cells }, { name: 'Q', cells: {
      A1: { value: 399 }, B1: { formula: 'HLOOKUP(A1,Data!1:2,2,FALSE)' },
      C1: { formula: 'MATCH(A1,Data!1:1,0)' }, D1: { formula: 'XLOOKUP(A1,Data!1:1,Data!2:2)' },
    } }] }, { clock, execution, display: false })
    try {
      for (const key of [399, 400, 300, 1]) {
        s.setCells([{ sheet: 'Q', cell: 'A1', input: { value: key } }])
        const v = s.cells('Q', 'B1:D1')
        expect([v.B1.value, v.C1.value, v.D1.value]).toEqual([key * 10, key, key * 10])
      }
      s.setCells([{ sheet: 'Data', cell: 'A2', input: { value: 123 } }])
      expect(s.cells('Q', 'B1').B1.value).toBe(123)
      expect(s.hash()).toBe(compute(s.book(), { clock, execution, display: false }).hash)
    } finally { s.close() }
  })

  it('follows a spill growing, blocked and released after the index was used', () => {
    const s = openSession({ sheets: [{ name: 'S', cells: {
      A1: { formula: 'SEQUENCE(C1)' }, C1: { value: 400 },
      E1: { value: 390 }, F1: { formula: 'XLOOKUP(E1,A1#,A1#)' }, G1: { formula: 'XMATCH(E1,A1#,0,-1)' },
    } }] }, { clock, execution, display: false })
    try {
      for (const key of [390, 399, 400]) {
        s.setCells([{ sheet: 'S', cell: 'E1', input: { value: key } }])
        expect(s.cells('S', 'F1').F1.value).toBe(key)
      }
      s.setCells([{ sheet: 'S', cell: 'C1', input: { value: 500 } }, { sheet: 'S', cell: 'E1', input: { value: 500 } }])
      expect([s.cells('S', 'F1').F1.value, s.cells('S', 'G1').G1.value]).toEqual([500, 500])
      s.setCells([{ sheet: 'S', cell: 'A250', input: { value: 9 } }])
      expect(s.cells('S', 'A1').A1.value).toEqual({ error: 'SPILL' })
      expect(s.hash()).toBe(compute(s.book(), { clock, execution, display: false }).hash)
      s.setCells([{ sheet: 'S', cell: 'A250', input: null }])
      expect(s.cells('S', 'F1').F1.value).toBe(500)
      expect(s.hash()).toBe(compute(s.book(), { clock, execution, display: false }).hash)
    } finally { s.close() }
  })

  it.each(['en-US', 'ru-RU'] as const)('preserves text equality, duplicate direction and wildcard semantics in %s', (locale) => {
    const cells: Record<string, CellInput> = {}
    for (let r = 1; r <= 400; r++) { cells[`A${r}`] = { value: `item ${r}` }; cells[`B${r}`] = { value: r } }
    const texts = ['Alpha', 'ALPHA', 'é', 'e\u0301', 'e', 'ё', 'е', 'й', 'и', 'a*b', 'a?b', 'a~b', '', 'ab']
    texts.forEach((value, i) => cells[`A${i + 1}`] = { value })
    const book: BookInput = { locale, sheets: [{ name: 'Data', cells }, { name: 'Q', cells: {
      A1: { value: 'Alpha' },
      B1: { formula: 'XLOOKUP(A1,Data!A:A,Data!B:B,"missing",0,1)' },
      C1: { formula: 'XLOOKUP(A1,Data!A:A,Data!B:B,"missing",0,-1)' },
      D1: { formula: 'XMATCH(A1,Data!A:A,0)' },
      E1: { formula: 'VLOOKUP(A1,Data!A:B,2,FALSE)' },
      F1: { formula: 'MATCH(A1,Data!A:A,0)' },
    } }] }
    const s = openSession(book, { clock, execution, display: false })
    try {
      // Misses warm the text indexes even when the following matches are near the start.
      for (let i = 0; i < 4; i++) s.setCells([{ sheet: 'Q', cell: 'A1', input: { value: `absent ${i}` } }])
      for (let repeat = 0; repeat < 2; repeat++) for (const [key, first, last] of [['alpha', 1, 2], ['É', 3, 4], ['e\u0301', 3, 4], ['e', 5, 5], ['Ё', 6, 6], ['е', 7, 7], ['Й', 8, 8], ['и', 9, 9], ['', 13, 13]] as const) {
        s.setCells([{ sheet: 'Q', cell: 'A1', input: { value: key } }])
        const v = s.cells('Q', 'B1:F1')
        expect([v.B1.value, v.C1.value, v.D1.value, v.E1.value, v.F1.value]).toEqual([first, last, first, first, first])
      }
      s.setCells([{ sheet: 'Q', cell: 'A1', input: { value: 'a~*b' } }])
      const v = s.cells('Q', 'B1:F1')
      expect([v.B1.value, v.C1.value, v.E1.value, v.F1.value]).toEqual(['missing', 'missing', 10, 10])
      expect(s.hash()).toBe(compute(s.book(), { clock, execution, display: false }).hash)
    } finally { s.close() }
  })

  it('refreshes after data edits, range growth and structure changes; a return cell edit keeps the key positions correct', () => {
    const cells: Record<string, CellInput> = {}
    for (let r = 1; r <= 600; r++) { cells[`A${r}`] = { value: r }; cells[`B${r}`] = { value: r * 10 } }
    const book: BookInput = { sheets: [{ name: 'Data', cells }, { name: 'Q', cells: {
      A1: { value: 599 }, B1: { formula: 'VLOOKUP(A1,Data!A:B,2,FALSE)' },
      C1: { formula: 'XMATCH(A1,Data!A:A,0,-1)' }, D1: { formula: 'XLOOKUP(A1,Data!A:A,Data!B:B)' },
    } }] }
    const s = openSession(book, { clock, execution, display: false })
    const check = (key: number, value: number, row: number) => {
      s.setCells([{ sheet: 'Q', cell: 'A1', input: { value: key } }])
      const v = s.cells('Q', 'B1:D1')
      expect([v.B1.value, v.C1.value, v.D1.value]).toEqual([value, row, value])
      expect(s.hash()).toBe(compute(s.book(), { clock, execution, display: false }).hash)
    }
    try {
      check(599, 5990, 599); check(600, 6000, 600)
      s.setCells([{ sheet: 'Data', cell: 'A599', input: { value: 900 } }]); check(900, 5990, 599)
      s.setCells([{ sheet: 'Data', cell: 'B599', input: { value: 123 } }]); check(900, 123, 599)
      s.setCells([{ sheet: 'Data', cell: 'A900', input: { value: 1000 } }, { sheet: 'Data', cell: 'B900', input: { value: 42 } }]); check(1000, 42, 900)
      expect(s.restructure({ op: 'insertRows', sheet: 'Data', row: 2, count: 2 }).ok).toBe(true); check(1000, 42, 902)
      expect(s.restructure({ op: 'deleteRows', sheet: 'Data', row: 2, count: 2 }).ok).toBe(true); check(900, 123, 599)
    } finally { s.close() }
  })
})
