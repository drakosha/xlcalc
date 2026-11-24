import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { DISPLAY_CASES, displayFile } from '../bench/coverage/table-styles/display-held.js'
import { REGION_CASES, regionFile } from '../bench/coverage/table-styles/regions-held.js'
import type { BorderLine } from '../src/contract.js'
import { libUrl } from './fixtures/target.js'

const { createSession, viewportOf } = await import(libUrl('index.js')) as typeof import('../src/index.js')
const { readXlsx } = await import(libUrl('xlsx.js')) as typeof import('../src/xlsx.js')
const { registerTableStyles } = await import(libUrl('table-styles.js')) as typeof import('../src/table-styles.js')
registerTableStyles()

type ExcelEdge = [lineStyle: number, weight: number, color: number]
type ExcelCell = [address: string, text: string, fill: number, pattern: number, color: number, bold: boolean, italic: boolean, sides: ExcelEdge[]]
type Answers = Record<string, [sheet: string, cells: ExcelCell[]][]>
const answersOf = (set: string): Answers => JSON.parse(readFileSync(new URL(`excel/win/xlsx/table-style-${set}.json`, import.meta.url), 'utf8')).answers
const displayAnswers = answersOf('display-held')
const regionAnswers = answersOf('regions-held')
const cases = [
  ...DISPLAY_CASES.map((c) => ({ ...c, file: () => displayFile(c.id), answers: displayAnswers[c.id] })),
  ...REGION_CASES.map((c) => ({ ...c, file: () => regionFile(c.id), answers: regionAnswers[c.id] })),
]
const colorOf = (windows: number) => '#' + [windows & 255, windows >> 8 & 255, windows >> 16 & 255]
  .map((part) => part.toString(16).padStart(2, '0')).join('').toUpperCase()
const excelEdge = ([style, weight, color]: ExcelEdge): BorderLine | undefined => {
  if (style === -4142) return undefined
  if (style === -4119) return { style: 'double', color: colorOf(color) }
  if (style === 1 && weight === 2) return { style: 'thin', color: colorOf(color) }
  throw new Error(`Unknown Excel border: ${style}, weight ${weight}`)
}

describe('Excel display styles in viewports', () => {
  it.each(cases)('matches the held Excel DisplayFormat for $id', async ({ id, sheet, range, file, answers }) => {
    const { book, cached } = await readXlsx(await file())
    const session = createSession(book, { cached })
    try {
      const viewport = session.getViewport(sheet, range)
      const saved = viewportOf(session.getSnapshot(), sheet, range, { appearance: true })
      expect(saved.resolvedStyles).toEqual(viewport.resolvedStyles)
      for (const [address, , fill, pattern, color, bold, italic, sides] of answers[0][1]) {
        const cell = viewport.cells[address]
        const style = cell?.resolvedStyle !== undefined ? viewport.resolvedStyles?.[cell.resolvedStyle] : undefined
        const where = `${sheet}!${address}`
        expect(style?.fill?.color, where).toBe(pattern === -4142 ? undefined : colorOf(fill))
        expect(style?.font?.color ?? '#000000', where).toBe(colorOf(color))
        expect(style?.font?.bold ?? false, where).toBe(bold)
        expect(style?.font?.italic ?? false, where).toBe(italic)
        expect((['left', 'top', 'bottom', 'right'] as const).map((side) => style?.border?.[side]), where).toEqual(sides.map(excelEdge))
      }
      if (id === 'banded-table') {
        expect(viewport.cells.C2.value).toBeNull()
        const styled = viewport.cells.B2
        expect(viewport.styles[styled.effectiveStyle!].fill?.color).toBe('#FFFF00')
        expect(styled.appearance?.style?.fill?.color).toBe('#00FF00')
      }
    } finally {
      session.close()
    }
  })
})

it('counts visible rows for table bands even when the window starts later in the table', async () => {
  const { book, cached } = await readXlsx(await displayFile('banded-table'))
  book.sheets[0].rows = { ...book.sheets[0].rows, 3: { hidden: true } }
  const session = createSession(book, { cached })
  try {
    const view = session.getViewport('S', 'C2:C5', { rows: [2, 4, 5] })
    const fill = (address: string) => view.resolvedStyles?.[view.cells[address].resolvedStyle!]?.fill?.color
    expect(fill('C2')).toBe('#C0E6F5')
    expect(fill('C4')).toBeUndefined()
  } finally {
    session.close()
  }
})

it('keeps custom stripe sizes after hiding a row and a column', async () => {
  const { book, cached } = await readXlsx(await regionFile('custom-bands'))
  book.sheets[0].rows = { 3: { hidden: true } }
  book.sheets[0].columns = { C: { hidden: true } }
  const session = createSession(book, { cached })
  try {
    const view = session.getViewport('S', 'B3:D6')
    const style = (address: string) => view.resolvedStyles?.[view.cells[address].resolvedStyle!]
    expect(style('B3')?.fill).toBeUndefined()
    expect(style('D4')?.fill?.color).toBe('#AAFF99')
    expect(style('D5')?.fill?.color).toBe('#AAFF99')
    expect(style('D6')?.fill?.color).toBe('#FFDD44')
    expect(style('D4')?.font).toEqual({ color: '#122113', italic: true })
  } finally {
    session.close()
  }
})
