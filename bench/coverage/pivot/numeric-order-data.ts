/** Fixed input vectors that distinguish the association of six numbers. */
import type { SourceColumn } from './data.js'

export function numericVectors(): number[][] {
  let seed = 0x6a09e667
  const rnd = () => ((seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0) / 2 ** 32)
  return Array.from({ length: 16 }, (_, k) => Array.from({ length: 6 }, (_, i) => {
    const scale = 10 ** ((k + i * 3) % 7 - 3)
    return Number(((rnd() * 2 - 0.9) * scale).toPrecision(12))
  }))
}

/** Row/column groups AX={a,b}, AY={c}, BX={d}, BY={e,f}; the record order is e,a,d,c,b,f. */
export const NUMERIC_ORDER = [4, 0, 3, 2, 1, 5]
export function numericOrderColumns(): SourceColumn[] {
  const v = numericVectors()[0]
  return [
    { name: 'R', values: ['B', 'A', 'B', 'A', 'A', 'B'].map(value => ({ value })) },
    { name: 'C', values: ['Y', 'X', 'X', 'Y', 'X', 'Y'].map(value => ({ value })) },
    { name: 'V', values: NUMERIC_ORDER.map(i => ({ value: v[i] })) },
  ]
}
