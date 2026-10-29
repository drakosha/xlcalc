/**
 * Paste Special between workbooks of different themes against both Excels (the paste set `theme`), which the engine
 * has no paste of: what a host gets that carries the cells over by reading both files (`carried`) has the colors All
 * Using Source Theme gives, the source's theme resolved, where All and Formats take the target's theme colors; and the
 * font names it carries are neither's: a font of the source's minor scheme, and one named as the source's default font,
 * come without a name, so the target's default font, where Source Theme names the source's font on every cell and All
 * keeps a font named without a scheme.
 */
import { existsSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { CellStyle } from '../src/index.js'
import { shownFile, type Shown } from '../bench/coverage/paste/shown.js'
import { buildPasteCases, carried } from '../bench/coverage/paste/theme.js'

type Cells = Record<string, { value?: unknown; style?: CellStyle }>

/** The pasted cells C1:C6 of a book shown. */
const pasted = (shown: Shown) => {
  const cells = shown.S.cells as Cells
  return ['C1', 'C2', 'C3', 'C4', 'C5', 'C6'].map((a) => cells[a])
}

/** A cell's style without its font's name. */
const unnamed = (c: Cells[string]) => {
  const { font: { name: _, ...font } = {}, ...rest } = c.style ?? {}
  return Object.keys(font).length ? { ...rest, font } : rest
}

/** What All and Formats give C1:C6: accent 1 and 2 of the target's theme, its minor font, the explicit color and the named font kept. */
const TARGET_THEME: CellStyle[] = [{ fill: { color: '#4472C4' } }, { font: { color: '#ED7D31' } }, { font: { bold: true } }, { fill: { color: '#00B050' } }, { fill: { color: '#8EA9DB' } }, { font: { name: 'Georgia' } }]

const cases = buildPasteCases().cases

describe.each(['paste', 'win/paste'].filter((dir) => existsSync(new URL(`./excel/${dir}/theme.json`, import.meta.url))))('the paste set theme in test/excel/%s', (dir) => {
  const saved = async (id: string) => pasted(await shownFile(new Uint8Array(readFileSync(new URL(`./excel/${dir}/theme/${id}.xlsx`, import.meta.url)))))
  it('All Using Source Theme gives the colors a host carries over, and names the source font on every cell', async () => {
    const [excel, host] = [await saved('allUsingSourceTheme'), pasted(await carried(cases.find((c) => c.id === 'allUsingSourceTheme')!))]
    expect(excel.map(unnamed)).toEqual(host.map(unnamed))
    expect(excel.map((c) => c.style?.font?.name)).toEqual(Array(6).fill('Georgia'))
    expect(host.map((c) => c.style?.font?.name)).toEqual(Array(6).fill(undefined))
  })
  it.each(['all', 'formats'])('%s gives the target theme colors and minor font, which a host does not carry over', async (id) => {
    const [excel, host] = [await saved(id), pasted(await carried(cases.find((c) => c.id === id)!))]
    expect(excel.map((c) => c.style)).toEqual(TARGET_THEME)
    expect(excel.map((c) => c.value)).toEqual(id === 'formats' ? Array(6).fill('old') : [1, 2, 'minor', 4, 5, 'named'])
    expect(host.map((c) => c.style)).not.toEqual(TARGET_THEME)
  })
})
