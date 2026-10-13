/**
 * `Session.getEnd` (Control and an arrow key, as Excel's `Range.End`, whose rules `end-win.test.ts` checks against
 * Excel) against a walk cell by cell over what the session's own result and book give: the cells with a value, the
 * whole range of each spill, the hidden rows and columns. Seeded books with values, empty texts, looks alone, spills
 * (one of a whole column) and hidden lines, through edits of cells, of the layout and of the structure and an undo, in
 * a synchronous session, in the same thread and in a thread. Arguments it refuses; its speed on the synthetic book.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { createSession, InputError, type AsyncSession, type BookInput, type CellInput, type ComputeResult, type EndDirection } from '../src/index.js'
import { formatAddress, indexToCol, MAX_COLS, MAX_ROWS, parseAddress } from '../src/model.js'
import { parseRange } from '../src/structure/area.js'
import { lines } from '../src/layout/spans.js'
import { isHiddenColumn, isHiddenRow } from '../src/layout/hidden.js'
import { mulberry32, synthBook } from '../src/synth.js'
import { CLOCK } from './async-books.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const DIRECTIONS: EndDirection[] = ['up', 'down', 'left', 'right']
/** The seeded part of a sheet: rows and columns past it hold nothing and none of them is hidden. */
const [ROWS, COLS] = [40, 12]

/** A sheet of values, empty texts, looks alone and spills over the seeded part, some of its lines hidden. */
function seeded(seed: number): BookInput {
  const rnd = mulberry32(seed)
  const cells: Record<string, CellInput> = {}
  for (let r = 0; r < ROWS; r++)
    for (let c = 0; c < COLS; c++) {
      const x = rnd()
      const at = formatAddress(r, c)
      if (x < 0.3) cells[at] = { value: Math.floor(rnd() * 9) }
      else if (x < 0.34) cells[at] = { formula: '""' }
      else if (x < 0.36) cells[at] = { value: '' }
      else if (x < 0.4) cells[at] = { style: 0 }
      else if (x < 0.41) cells[at] = { formula: `SEQUENCE(${1 + Math.floor(rnd() * 4)},${1 + Math.floor(rnd() * 3)})` }
    }
  const hidden = (n: number, key: (i: number) => string) => Object.fromEntries(Array.from({ length: n }, (_, i) => i).filter(() => rnd() < 0.15).map((i) => [key(i), { hidden: true }]))
  return {
    styles: [{ font: { bold: true } }],
    sheets: [{ name: 'S', cells, rows: hidden(ROWS, (i) => String(i + 1)), columns: hidden(COLS, indexToCol) }],
  }
}

/** The cell `End` comes to by the rule, walking line by line over the result's cells and the book's hidden lines. */
function walked(result: ComputeResult, book: BookInput, sheet: string, cell: string, d: EndDirection): string {
  const held = new Set<string>()
  for (const [a, c] of Object.entries(result.sheets.find((x) => x.name === sheet)!.cells)) {
    if (c.value !== null) held.add(a)
    if (!c.spill) continue
    const x = parseRange(c.spill)!
    for (let r = x.r1; r <= x.r2; r++) for (let k = x.c1; k <= x.c2; k++) held.add(formatAddress(r, k))
  }
  const input = book.sheets.find((x) => x.name === sheet)!
  const hiddenOf = <T,>(map: Record<string, T> | undefined, axis: 'r' | 'c', test: (e: T) => boolean) => new Set(lines(map, axis).flatMap(([lo, hi, e]) => (test(e) ? Array.from({ length: hi - lo + 1 }, (_, i) => lo + i) : [])))
  const vertical = d === 'up' || d === 'down'
  const step = d === 'down' || d === 'right' ? 1 : -1
  const hidden = vertical ? hiddenOf(input.rows, 'r', isHiddenRow) : hiddenOf(input.columns, 'c', isHiddenColumn)
  const max = (vertical ? MAX_ROWS : MAX_COLS) - 1
  const { row, col } = parseAddress(cell)
  const at = (i: number) => (vertical ? formatAddress(i, col) : formatAddress(row, i))
  // past the seeded part nothing is held or hidden: a walk past it goes to the edge
  const far = vertical ? ROWS + 6 : COLS + 4
  const next = (i: number) => {
    if (step > 0 && i >= far) return i === max ? -1 : max
    if (step < 0 && i > far) return far
    for (i += step; i >= 0 && i <= max && hidden.has(i); i += step);
    return i < 0 || i > max ? -1 : i
  }
  const start = vertical ? row : col
  let i = next(start)
  if (i < 0) return cell
  if (held.has(at(start)) && held.has(at(i))) {
    for (let j = next(i); j >= 0 && held.has(at(j)); j = next(j)) i = j
    return at(i)
  }
  for (let last = i; i >= 0; i = next(i)) {
    if (held.has(at(i))) return at(i)
    last = i
    if (next(i) < 0) return at(last)
  }
  return cell
}

/** Cells to start from: the seeded part, its edge and past it, and the sheet's edges. */
function starts(rnd: () => number, n: number): string[] {
  const pick = (k: number) => (rnd() < 0.85 ? Math.floor(rnd() * k) : rnd() < 0.5 ? k + Math.floor(rnd() * 3) : 0)
  const out = ['A1', 'XFD1', 'A1048576', 'XFD1048576', formatAddress(ROWS + 2, 2), formatAddress(5, COLS + 2)]
  while (out.length < n) out.push(formatAddress(rnd() < 0.05 ? MAX_ROWS - 1 : pick(ROWS + 2), rnd() < 0.05 ? MAX_COLS - 1 : pick(COLS + 2)))
  return out
}

/** Every start and direction as the walk gives it. */
async function check(s: AsyncSession, rnd: () => number, at: string) {
  const [result, book] = [await s.getResult(), await s.getBook()]
  const off: unknown[] = []
  for (const cell of starts(rnd, 30))
    for (const d of DIRECTIONS) {
      const [got, want] = [await s.getEnd('S', cell, d), walked(result, book, 'S', cell, d)]
      if (got !== want) off.push({ at, cell, d, got, want })
    }
  expect(off).toEqual([])
}

/** A spill down a whole column (B1:B1048576, its cells past the third held as one value), and cells beside it. */
const FAR: BookInput = { sheets: [{ name: 'Far', cells: { A1: { value: 1 }, A3: { value: 2 }, B1: { formula: 'A:A*2' }, D5: { value: 3 } } }] }
/** `cell direction` → where `End` comes to, with the spill and once B9 blocks it. */
const FAR_ENDS: Record<string, [spilled: string, blocked: string]> = {
  'B1 down': ['B1048576', 'B9'],
  'B2 down': ['B1048576', 'B9'],
  'B9 down': ['B1048576', 'B1048576'],
  'B1048576 up': ['B1', 'B9'],
  'B500000 up': ['B1', 'B9'],
  'A1 right': ['B1', 'B1'],
  'C7 left': ['B7', 'A7'],
  'XFD4 left': ['B4', 'A4'],
  'D5 left': ['B5', 'A5'],
  'A1048576 right': ['B1048576', 'XFD1048576'],
  'A1048576 up': ['A3', 'A3'],
}

describe.each(KINDS)('getEnd, %s', (kind) => {
  it('seeded books: each cell and direction as the walk over the result and the book gives it, through edits and an undo', async () => {
    const open = await opener(kind)
    for (const seed of [1, 2, 3]) {
      const rnd = mulberry32(100 + seed)
      const s = await open(seeded(seed), { clock: CLOCK })
      await check(s, rnd, `seed ${seed}`)
      await s.setCells([{ sheet: 'S', cell: 'B3', input: { formula: 'SEQUENCE(6)' } }, { sheet: 'S', cell: 'D1', input: null }, { sheet: 'S', cell: 'E9', input: { value: 'x' } }])
      await check(s, rnd, 'cells')
      await s.setLayout([{ sheet: 'S', rows: '2:4', set: { hidden: true } }, { sheet: 'S', columns: 'C', set: { hidden: false } }, { sheet: 'S', columns: 'F:G', set: { hidden: true } }])
      await check(s, rnd, 'layout')
      await s.editStructure({ op: 'insertRows', sheet: 'S', row: 3, count: 2 })
      await check(s, rnd, 'rows inserted')
      await s.editStructure({ op: 'deleteColumns', sheet: 'S', column: 'B', count: 2 })
      await check(s, rnd, 'columns deleted')
      await s.undo()
      await check(s, rnd, 'undone')
      await s.close()
    }
  }, 120_000)

  it('a spill down a whole column: its cells to the edge, and as far as a value that blocks it', async () => {
    const s = await (await opener(kind))(FAR, { clock: CLOCK })
    const ends = async () => Object.fromEntries(await Promise.all(Object.keys(FAR_ENDS).map(async (k) => [k, await s.getEnd('Far', ...(k.split(' ') as [string, EndDirection]))])))
    expect(await ends()).toEqual(Object.fromEntries(Object.entries(FAR_ENDS).map(([k, v]) => [k, v[0]])))
    await s.setCells([{ sheet: 'Far', cell: 'B9', input: { value: 0 } }])
    expect(await ends()).toEqual(Object.fromEntries(Object.entries(FAR_ENDS).map(([k, v]) => [k, v[1]])))
    await s.close()
  })

  it('refuses a sheet, a cell or a direction it does not know', async () => {
    const s = await (await opener(kind))(seeded(1), { clock: CLOCK })
    await expect(s.getEnd('Nope', 'A1', 'down')).rejects.toThrow(InputError)
    await expect(s.getEnd('S', 'A0', 'down')).rejects.toThrow(InputError)
    await expect(s.getEnd('S', 'A1', 'north' as EndDirection)).rejects.toThrow(new InputError('direction', 'expected up, down, left or right, got the string "north"'))
    // the sheet's name in any case, as every call takes it
    expect(await s.getEnd('s', 'A1048576', 'up')).toBe(await s.getEnd('S', 'A1048576', 'up'))
    await s.close()
  })
})

describe('getEnd on the synthetic book of 34,000 rows', () => {
  it('under 1 ms a call in a synchronous session, down a whole column of values and across the sheet', () => {
    const { book } = synthBook({ seed: 1, rows: 34_000 })
    const s = createSession(book, { clock: CLOCK })
    const sheet = book.sheets[0].name
    const calls: [string, EndDirection, string][] = [['A1', 'down', 'A34001'], ['A34001', 'up', 'A1'], ['C17000', 'down', 'C34001'], ['A1048576', 'up', 'A34001'], ['A1', 'right', 'O1'], ['XFD17000', 'left', 'O17000']]
    for (const [cell, d, want] of calls) {
      expect(s.getEnd(sheet, cell, d), `${cell} ${d}`).toBe(want)
      const times = Array.from({ length: 21 }, () => {
        const t = performance.now()
        s.getEnd(sheet, cell, d)
        return performance.now() - t
      }).sort((a, b) => a - b)
      expect(times[10], `${cell} ${d}`).toBeLessThan(1)
    }
    s.close()
  }, 60_000)
})
