/**
 * Custom table styles' inside borders, patterned fills and number formats in a viewport, against Excel for Windows's
 * display format of the same workbook (`bench/windows/table-styles-inside.ts`, prediction in
 * `bench/coverage/table-styles/inside-held.predicted.json`).
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { INSIDE_CASES, insideFile } from '../bench/coverage/table-styles/inside-held.js'
import type { BorderLine, FillPattern, FillStyle } from '../src/contract.js'
import { libUrl } from './fixtures/target.js'

const { createSession } = await import(libUrl('index.js')) as typeof import('../src/index.js')
await import(libUrl('locales/ru-RU.js'))
const { readXlsx, writeXlsx } = await import(libUrl('xlsx.js')) as typeof import('../src/xlsx.js')
const { registerTableStyles } = await import(libUrl('table-styles.js')) as typeof import('../src/table-styles.js')
registerTableStyles()

type ExcelEdge = [lineStyle: number, weight: number, color: number]
type ExcelCell = [address: string, text: string, fill: number, colorIndex: number, font: number, bold: boolean, italic: boolean, edges: ExcelEdge[], pattern: number, patternColor: number, format: string]
const answers: Record<string, [string, ExcelCell[]][]> = JSON.parse(readFileSync(new URL('excel/win/xlsx/table-style-inside-held.json', import.meta.url), 'utf8')).answers

const colorOf = (bgr: number) => '#' + [bgr & 255, bgr >> 8 & 255, bgr >> 16 & 255].map((part) => part.toString(16).padStart(2, '0')).join('').toUpperCase()
// XlPattern as Excel's object model gives it, by xlsx name
const PATTERNS: Record<number, FillPattern> = { 9: 'darkGrid', 10: 'darkTrellis', 14: 'lightUp', 17: 'gray125' }
const LINES: Record<string, BorderLine['style']> = { '1,2': 'thin', '-4115,2': 'dashed', '1,-4138': 'medium' }

function excelFill(fill: number, pattern: number, patternColor: number): FillStyle | undefined {
  if (pattern === -4142) return undefined
  if (pattern === 1) return { color: colorOf(fill) }
  const name = PATTERNS[pattern]
  if (!name) throw new Error(`Unknown Excel pattern ${pattern}`)
  // 0 is Excel's automatic pattern colour, which the model leaves out
  return { color: colorOf(fill), pattern: name, ...(patternColor ? { patternColor: colorOf(patternColor) } : {}) }
}

function excelEdge([style, weight, color]: ExcelEdge): BorderLine | undefined {
  if (style === -4142) return undefined
  const kind = LINES[`${style},${weight}`]
  if (!kind) throw new Error(`Unknown Excel border: ${style}, weight ${weight}`)
  return { style: kind, color: colorOf(color) }
}

const spaces = (text: string) => text.replace(/[  ]/g, ' ')

describe('custom table styles beyond a cell style, against Excel', () => {
  it.each(INSIDE_CASES)('$id: inside borders, patterns and number formats as Excel shows them', async ({ id, sheet, range }) => {
    const { book, cached } = await readXlsx(await insideFile(id))
    const session = createSession({ ...book, locale: 'ru-RU' }, { cached })
    try {
      const view = session.getViewport(sheet, range)
      for (const [address, text, fill, , , bold, , edges, pattern, patternColor] of answers[id][0][1]) {
        const cell = view.cells[address]
        const style = cell?.resolvedStyle === undefined ? undefined : view.resolvedStyles?.[cell.resolvedStyle]
        const where = `${sheet}!${address}`
        expect(spaces(cell?.text ?? ''), where).toBe(spaces(text))
        expect(style?.fill, where).toEqual(excelFill(fill, pattern, patternColor))
        expect(style?.font?.bold ?? false, where).toBe(bold)
        expect((['left', 'top', 'bottom', 'right'] as const).map((side) => style?.border?.[side]), where).toEqual(edges.map(excelEdge))
      }
    } finally {
      session.close()
    }
  })

  it('writes the definitions back as they were read', async () => {
    const { book } = await readXlsx(await insideFile('inside'))
    const again = await readXlsx((await writeXlsx({ book })).data)
    // a built-in code reads back as its number, as a rule's dxf does: `0.00%` is 10
    expect(again.book.tableStyles!.custom!.Probe.elements.find((e) => e.type === 'totalRow')?.numberFormat).toBe(10)
    const third = await readXlsx((await writeXlsx({ book: again.book })).data)
    expect(third.book.tableStyles).toEqual(again.book.tableStyles)
    const probe = book.tableStyles!.custom!.Probe.elements
    expect(probe.find((e) => e.type === 'wholeTable')?.inside).toEqual({ horizontal: { style: 'thin', color: '#FF0000' }, vertical: { style: 'dashed', color: '#0000FF' } })
    expect(probe.find((e) => e.type === 'headerRow')?.style.fill).toEqual({ color: '#FFFF00', pattern: 'lightUp', patternColor: '#FF0000' })
    expect(probe.find((e) => e.type === 'totalRow')?.numberFormat).toBe('0.00%')
  })
})
