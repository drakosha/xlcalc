import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { bookGeometry, columnPoints } from '../src/column-width.js'
import { DEFAULT_WIDTH_CASES, defaultWidthFile, defaultWidthPrediction } from '../bench/coverage/xlsx/default-column-width.js'

import { createSession } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { writeKept } from '../src/xlsx-keep.js'
import { attrsAt } from '../src/xlsx/xml/scan.js'
import { openArchive } from '../src/xlsx/zip/read.js'

const { answers } = JSON.parse(readFileSync(new URL('excel/xlsx/default-column-width.json', import.meta.url), 'utf8')) as {
  answers: Record<string, {
    columns: [points: number, characters: number, shown: string][]
    saved: { defaultColWidth: string; defaultRowHeight: string }
  }>
}

describe('the default column width of a sheet', () => {
  it.each(DEFAULT_WIDTH_CASES)('$id has the widths Excel for Mac shows', async (c) => {
    const result = await defaultWidthPrediction(c)
    const { columns } = answers[c.id]
    expect(result.points).toEqual(columns.map(([points]) => points))
    expect(result.shown).toEqual(columns.map(([, , shown]) => shown.replace(',', '.')))
    expect(result.issues).toEqual([])
  })

  it('writes the default width and uses it for a column with formatting alone', async () => {
    const { book } = await readXlsx(await defaultWidthFile({ id: 'narrow', defaultColWidth: 4.5 }))
    book.sheets[0].columns!.C = { numberFormat: '0.0' }
    const { data } = await writeXlsx({ book })
    const archive = openArchive(data)
    const xml = (await archive.text('xl/worksheets/sheet1.xml'))!
    expect(attrsAt(xml, 'sheetFormatPr').defaultColWidth).toBe('4.5')
    expect(xml).toContain('<col min="3" max="3" width="4.5" style=')
    const back = await readXlsx(data)
    expect(back.book.sheets[0].defaultColumnWidth).toBe(4.5)
    expect(back.book.sheets[0].columns?.C).toEqual({ numberFormat: '0.0' })
  })

  it.each(DEFAULT_WIDTH_CASES.filter(({ font }) => font))('$id writes the inherited width in the Normal font', async (c) => {
    const original = await defaultWidthFile(c)
    const { book } = await readXlsx(original)
    const { saved, columns: [[points]] } = answers[c.id]
    const [digit] = bookGeometry(book.font)
    const originalWidth = book.sheets[0].defaultColumnWidth
    expect(book.font).toEqual(c.font)
    expect(columnPoints(originalWidth, digit)).toBe(points)
    if (c.baseColWidth !== undefined) expect(originalWidth).toBe(Number(saved.defaultColWidth))

    book.sheets[0].columns!.C = { numberFormat: '0.0' }
    for (const { data } of [await writeXlsx({ book }), await writeKept(original, { book })]) {
      const xml = (await openArchive(data).text('xl/worksheets/sheet1.xml'))!
      expect(attrsAt(xml, 'sheetFormatPr').defaultRowHeight).toBe(saved.defaultRowHeight)
      expect(xml).toContain(`<col min="3" max="3" width="${originalWidth}" style=`)
      const back = await readXlsx(data)
      expect(back.book.sheets[0].defaultColumnWidth).toBe(originalWidth)
      expect(back.book.sheets[0].columns?.C).toEqual({ numberFormat: '0.0' })
    }
  })

  it('updates and clears the width in a kept file that specified a base width', async () => {
    const original = await defaultWidthFile({ id: 'base', baseColWidth: 20 })
    const { book } = await readXlsx(original)
    expect(book.sheets[0].defaultColumnWidth).toBe(20.83203125)
    book.sheets[0].defaultColumnWidth = 10
    const written = await writeKept(original, { book })
    const changed = await readXlsx(written.data)
    expect(changed.book.sheets[0].defaultColumnWidth).toBe(10)

    delete changed.book.sheets[0].defaultColumnWidth
    const cleared = await writeKept(written.data, { book: changed.book })
    expect((await readXlsx(cleared.data)).book.sheets[0].defaultColumnWidth).toBeUndefined()
  })

  it('uses the width for displayed numbers and retains it during column insertion', async () => {
    const { book } = await readXlsx(await defaultWidthFile({ id: 'narrow', defaultColWidth: 4.5 }))
    const session = createSession(book)
    try {
      expect(session.getCell('S', 'A1')?.text).toBe(answers.narrow.columns[0][2])
      expect(session.getCell('S', 'B1')?.text).toBe('12345.67')
      expect(session.editStructure({ op: 'insertColumns', sheet: 'S', column: 'A', count: 1 }).ok).toBe(true)
      expect(session.getLayout().sheets[0].defaultColumnWidth).toBe(4.5)
      expect(session.getViewport('S', 'A1:C1').columns.map(({ width }) => width)).toEqual([4.5, 4.5, 12])
    } finally {
      session.close()
    }
  })
})
