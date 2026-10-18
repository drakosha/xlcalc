/**
 * What SUMIFS costs when it selects again. A data sheet of 3,000 rows: A one of two directions, B one of 100 items (30
 * rows each), C amounts; 50 formulas sum an item's income. After the book opened and one edit, an edit of the directions
 * lowers each distinct text of the criteria columns and criteria at most once, not each candidate a criterion checks.
 */
import { describe, expect, it } from 'vitest'
import { compute, createSession, type BookInput, type CellInput } from '../src/index.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const DIRS = ['Доход', 'Расход']
const ITEMS = Array.from({ length: 100 }, (_, i) => `Товар ${'АБВГДЕЖЗИК'[Math.floor(i / 10)]}${'абвгдежзик'[i % 10]}`)

/** The data sheet S with 50 formulas in E, each `SUMIFS(C:C,<the item>,<"Доход">)` with the item's criterion first or the direction's. */
function book(itemFirst: boolean): BookInput {
  const cells: Record<string, CellInput> = {}
  for (let r = 1; r <= 3000; r++) {
    cells[`A${r}`] = { value: DIRS[r % 2] }
    cells[`B${r}`] = { value: ITEMS[r % 100] }
    cells[`C${r}`] = { value: r % 97 }
  }
  for (let k = 1; k <= 50; k++) {
    const [item, dir] = [`B:B,"${ITEMS[k]}"`, `A:A,"${DIRS[0]}"`]
    cells[`E${k}`] = { formula: `SUMIFS(C:C,${itemFirst ? `${item},${dir}` : `${dir},${item}`})` }
  }
  return { sheets: [{ name: 'S', cells }] }
}

/** The calls of `String.prototype.toLowerCase` while `f` runs. */
function lowered(f: () => void): number {
  const own = String.prototype.toLowerCase
  let n = 0
  String.prototype.toLowerCase = function (this: string) {
    n++
    return own.call(this)
  }
  try {
    f()
  } finally {
    String.prototype.toLowerCase = own
  }
  return n
}

describe.each(['closures', 'eval'] as const)('criteria, %s', (execution) => {
  const opts = { clock, execution, display: false }

  it('an edit of a criteria column lowers each distinct text once, not each candidate', () => {
    const s = createSession(book(true), opts)
    s.setCells([{ sheet: 'S', cell: 'A2', input: { value: DIRS[1] } }])
    const n = lowered(() => s.setCells([{ sheet: 'S', cell: 'A4', input: { value: DIRS[1] } }]))
    // the texts of the directions and items, and of the criteria: two directions, 100 items, 51 criteria
    expect(n).toBeLessThanOrEqual(DIRS.length + ITEMS.length + 51)
    expect(s.getHash()).toBe(compute(s.getBook(), opts).hash)
  })
})
