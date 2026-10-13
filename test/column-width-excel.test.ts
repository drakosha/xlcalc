/** Object-model points, file widths and defaults from the held-out probe on both platforms. */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { COLUMNS, FILE_WIDTHS } from '../bench/coverage/width/cases.js'
import { columnWidthToPoints } from '../src/text.js'
interface Measured {
  standardWidth: number
  columns: [characters: number, points: number, standard: boolean][]
  saved: Record<string, { width: number }>
}
const read = (path: string): Measured => JSON.parse(readFileSync(new URL(path, import.meta.url), 'utf8')) as Measured
const mac = read('./excel/width/width-held.json')
const win = read('./excel/win/width/width-held.json')

it('the platform defaults have different characters and different point widths', () => {
  expect([mac.standardWidth, mac.columns[0][1], win.standardWidth, win.columns[0][1]]).toEqual([8, 53, 8.43, 48])
})

describe('Mac width model', () => {
  it.each(mac.columns.map((c, i) => [COLUMNS[i], c] as const))('%s: the saved width restores the object-model points', (column, measured) => {
    expect(columnWidthToPoints(mac.saved[column]?.width)).toBe(measured[1])
  })
  it.each(FILE_WIDTHS.map((w, i) => [w, mac.columns[i + 1][1]]))('%s: the input file width also gives the points', (width, points) => {
    expect(columnWidthToPoints(width)).toBe(points)
  })
})

describe('Windows file-width formula', () => {
  it.each(win.columns.map((c, i) => [COLUMNS[i], c] as const))('%s: the saved width gives points with a 7-pixel digit', (column, measured) => {
    const width = win.saved[column]?.width
    const points = width === undefined ? 48 : Math.trunc((256 * width + Math.trunc(128 / 7)) * 7 / 256) * 3 / 4
    expect(points).toBe(measured[1])
  })
  it('WIN_KNOWN: the model keeps the Mac point geometry, so its default is not Windows 48 points', () => {
    expect(columnWidthToPoints()).toBe(53)
    expect(columnWidthToPoints()).not.toBe(win.columns[0][1])
    expect([columnWidthToPoints(8.83203125), win.columns[1][1]]).toEqual([53, 46.5])
  })
})
