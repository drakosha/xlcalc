/**
 * A column's width in characters as Excel for Mac's Column Width box shows and takes it (`columnWidthToChars`,
 * `columnWidthFromChars`), against every column the fit sets read on the Mac (`test/excel/fit/`, `ColumnWidth` beside
 * the stored width), both ways; and the sizes of rows and outlines.
 */
import { readdirSync, readFileSync } from 'node:fs'
import { expect, it } from 'vitest'
import { InputError } from '../src/index.js'
import { columnWidthFromChars, columnWidthToChars, DEFAULT_ROW_HEIGHT, MAX_OUTLINE, MAX_ROW_HEIGHT } from '../src/text.js'

type Answer = { columnWidth?: number; saved?: { width?: number } | null; standard?: boolean }

function measured(): [number | undefined, number][] {
  const out = new Map<string, [number | undefined, number]>()
  for (const dir of ['fit', 'fit/ru-RU']) {
    const at = new URL(`./excel/${dir}/`, import.meta.url)
    for (const f of readdirSync(at).filter((x) => x.endsWith('.json'))) {
      const answers = (JSON.parse(readFileSync(new URL(f, at), 'utf8')) as { answers: Record<string, Answer> }).answers
      for (const a of Object.values(answers)) {
        if (a.columnWidth === undefined) continue
        const w = a.standard ? undefined : a.saved?.width
        if (a.standard || w !== undefined) out.set(`${w}`, [w, a.columnWidth])
      }
    }
  }
  return [...out.values()]
}

it("every width the Mac's fit sets read shows Excel's characters, and those characters give it back", () => {
  const pairs = measured()
  expect(pairs.length).toBeGreaterThan(100)
  for (const [width, chars] of pairs) {
    expect([width, columnWidthToChars(width)]).toEqual([width, chars])
    if (width !== undefined) expect([chars, columnWidthFromChars(chars)]).toEqual([chars, width])
  }
  // the default column is 8 characters; 8 typed into the box are a width of its own, 53 points as the default
  expect([columnWidthToChars(), columnWidthToChars(null), columnWidthFromChars(8)]).toEqual([8, 8, 8.83203125])
})

it('the input checked; the sizes of rows and outlines', () => {
  expect(() => columnWidthFromChars(256)).toThrow(InputError)
  expect(() => columnWidthToChars(-1)).toThrow(InputError)
  expect([DEFAULT_ROW_HEIGHT, MAX_ROW_HEIGHT, MAX_OUTLINE]).toEqual([15, 409.5, 7])
})
