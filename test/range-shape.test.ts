import { describe, expect, it } from 'vitest'
import { compute, openSession } from '../src/index.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

describe.each(['closures', 'eval'] as const)('cached range shape, %s', (execution) => {
  it('keeps a cached rectangle aligned after a write outside it grows the sheet width', () => {
    const opts = { clock, execution, display: false }
    const s = openSession({ sheets: [
      { name: 'Data', cells: { A1: { value: 1 }, A2: { value: 2 } } },
      // the range is read twice, so its values are cached from the first computation on
      { name: 'Q', cells: { A1: { formula: 'INDEX(Data!A1:D256*1,2,1)+COUNTA(Data!A1:D256)*0+B1' }, B1: { value: 0 } } },
    ] }, opts)
    try {
      expect(s.cells('Q', 'A1').A1.value).toBe(2)
      s.setCells([{ sheet: 'Data', cell: 'E1000', input: { value: 42 } }])
      s.setCells([{ sheet: 'Q', cell: 'B1', input: { value: 1 } }])
      expect(s.cells('Q', 'A1').A1.value).toBe(3)
      expect(s.hash()).toBe(compute(s.book(), opts).hash)
    } finally { s.close() }
  })

  it('restores a whole-column spill after an obstruction grew the sheet height', () => {
    const opts = { clock, execution, display: false }
    const s = openSession({ sheets: [{ name: 'S', cells: {
      A1: { value: 1 }, A2: { value: 2 }, A257: { value: 13 },
      B1: { formula: 'A:A*2' }, C1: { formula: 'COUNT(B:B)' },
      C2: { formula: 'SUM(B1:B300)' }, C3: { formula: 'MAX(B:B)' }, C4: { formula: 'MIN(B:B)' },
    } }] }, opts)
    try {
      s.setCells([{ sheet: 'S', cell: 'B500', input: { value: 'block' } }])
      s.setCells([{ sheet: 'S', cell: 'B500', input: null }])
      expect(Object.values(s.cells('S', 'C1:C4')).map((c) => c.value)).toEqual([1048576, 32, 26, 0])
      const full = compute(s.book(), opts)
      expect(s.hash()).toBe(full.hash)
      expect(s.result().sheets).toEqual(full.sheets)
    } finally { s.close() }
  })
})
