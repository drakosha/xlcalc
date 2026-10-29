/** Engine-free column geometry uses the same Mac point model as formatting and its recorded oracle. */
import { readFileSync, readdirSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import * as text from '../src/text.js'
import { InputError } from '../src/index.js'
import { columnPoints, MAX_POINTS } from '../src/column-width.js'

it('gives the library default in points and restores the stored width after a drag', () => {
  expect([text.columnWidthToPoints(), text.columnWidthToPoints(null), text.columnWidthToPoints(8.83203125)]).toEqual([53, 53, 53])
  expect(text.columnWidthFromPoints(53)).toBe(8.83203125)
  expect(text.columnWidthToPoints(9.140625)).toBe(55)
})

it('keeps MAX_POINTS, a literal, the points of the widest stored width', () => {
  expect(MAX_POINTS).toBe(columnPoints(255.83203125))
  expect(() => text.columnWidthToPoints(255.83203125 + 1 / 256)).toThrow(InputError)
})

it('quantizes down to whole points and stores the width in 256ths of a digit', () => {
  for (let points = 0; points <= 1535; points++) {
    const width = text.columnWidthFromPoints(points)
    expect(width * 256).toBe(Math.floor(width * 256))
    expect(text.columnWidthToPoints(width)).toBe(points)
  }
  expect(text.columnWidthToPoints(text.columnWidthFromPoints(53.75))).toBe(53)
})

it('refuses wrong types, non-finite inputs and out-of-range widths', () => {
  for (const width of ['8', -1, 256, Infinity, NaN]) expect(() => text.columnWidthToPoints(width as never)).toThrow(InputError)
  for (const points of [undefined, null, '53', -1, 1536, Infinity, NaN]) expect(() => text.columnWidthFromPoints(points as never)).toThrow(InputError)
})

const dir = new URL('./excel/fit/', import.meta.url)
for (const file of readdirSync(dir).filter((f) => f.endsWith('.json'))) {
  const stored = JSON.parse(readFileSync(new URL(file, dir), 'utf8')) as { answers: Record<string, { points: number; saved: { width: number } | null }> }
  describe(`Mac point geometry, ${file}`, () => {
    it.each(Object.entries(stored.answers))('%s: the file width gives the object-model points', (_, answer) => {
      expect(text.columnWidthToPoints(answer.saved?.width)).toBe(answer.points)
    })
  })
}
