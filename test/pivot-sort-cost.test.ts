import { expect, it } from 'vitest'
import { sortPivotItems } from '../src/pivot/sort.js'

for (const descending of [false, true])
  for (const outlier of [false, true])
    it(`bounds work for a large tie group (${descending}, ${outlier})`, () => {
      const items = Array.from({ length: 5000 }, (_, i) => i)
      const key = (i: number) => outlier && i === 0 ? 1 : 0
      let comparisons = 0
      const compare = (a: number, b: number) => { comparisons++; return (key(a) - key(b)) * (descending ? -1 : 1) }
      sortPivotItems(items, (a, b) => { const c = compare(a, b); return c > 0 || (descending && c === 0) }, 1, compare)
      expect(comparisons).toBeLessThan(items.length * 20)
      if (outlier) expect(items[descending ? 0 : items.length - 1]).toBe(0)
    })
