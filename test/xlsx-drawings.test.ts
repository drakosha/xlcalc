import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { createSession } from '../src/index.js'
import { readXlsxCharts, resolveChartData } from '../src/xlsx-charts.js'
import { layoutPrintPages, printSheetGeometry, projectPrintDrawings, readXlsxDrawings } from '../src/xlsx-print.js'

import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { printingSnapshot } from './print-fixtures.js'

const kept = (name: string) => readFileSync(new URL('./excel/keep/open/' + name + '.xlsx', import.meta.url))

describe('XLSX drawings', () => {
  it('returns an Excel-saved picture, its anchor and lazily read image bytes', async () => {
    const file = await readXlsxDrawings(kept('picture-values'))
    expect(file.issues).toEqual([])
    expect(file.drawings).toContainEqual(expect.objectContaining({
      kind: 'picture', name: 'Picture 2', part: 'xl/media/image1.png', contentType: 'image/png',
      anchor: { kind: 'twoCell', from: { col: 6, row: 15, x: 35, y: 10 }, to: { col: 8, row: 19, x: 9, y: 0 }, editAs: 'oneCell' },
    }))
    const [image, samePart] = await Promise.all([file.readImage('xl/media/image1.png'), file.readImage('XL/MEDIA/IMAGE1.PNG')])
    expect(Array.from(image!.slice(0, 8))).toEqual([137, 80, 78, 71, 13, 10, 26, 10])
    expect(samePart).toBe(image)
    expect(await file.readImage('xl/workbook.xml')).toBeUndefined()
  })

  it.each([
    ['picture-values', ['E14:E20', 'F14:F20', 'G14:G20', 'H14:H20', 'I14:I20'], [0, 1, 1, 1, 0]],
    ['chart-values', ['E1:G5', 'H1:M5', 'N1:P5'], [1, 1, 0]],
  ] as const)('assigns %s objects to the pages where Windows Excel printed them', async (name, areas, expected) => {
    const data = kept(name)
    const [{ book }, { drawings }] = await Promise.all([readXlsx(data), readXlsxDrawings(data)])
    const dataSheet = book.sheets.find((sheet) => sheet.name === 'Data')!
    const geometry = printSheetGeometry(dataSheet, { font: book.font })
    const target = { width: 841.92, height: 595.32 }
    const { pages } = layoutPrintPages({ sheet: dataSheet.name, areas }, geometry, target)
    const objects = drawings.filter((drawing) => drawing.sheet === dataSheet.name)
    const placed = pages.map((page) => projectPrintDrawings(objects, page, geometry))
    expect(placed.map((list) => list.length)).toEqual(expected)
    for (const drawing of placed.flat()) {
      expect(drawing.regions[0].source.width).toBeLessThan(drawing.bounds.width)
    }
  })

  it('places a shown legacy note where its box intersects the page, even when its cell is outside', async () => {
    const settings = {
      sheet: 'S', areas: ['A1:A4', 'B1:B4', 'C1:C4'],
      setup: { cellComments: 'asDisplayed' as const },
    }
    const { data } = await writeXlsx(printingSnapshot({ book: { sheets: [{ name: 'S', cells: {
      A2: { note: { text: 'On the neighboring page', shown: true } },
      B2: { note: { text: 'Hidden' } },
    } }] } }, [settings]))
    const { drawings, issues } = await readXlsxDrawings(data)
    expect(issues).toEqual([])
    expect(drawings.map((drawing) => drawing.kind === 'note' && [drawing.cell, drawing.hidden])).toEqual([['A2', false], ['B2', true]])
    const geometry = printSheetGeometry({}, { usedRange: 'A1:C4' })
    const pages = layoutPrintPages(settings, geometry, { width: 400, height: 260 }).pages
    expect(pages.map((page) => projectPrintDrawings(drawings, page, geometry, { setup: settings.setup }).map((drawing) => drawing.drawing.id)))
      .toEqual([[], [drawings[0].id], [drawings[0].id]])
    const last = projectPrintDrawings(drawings, pages[2], geometry, { setup: settings.setup })[0]
    expect(last.regions[0].target.x).toBeLessThan(pages[2].body.x)
    expect(last.regions[0].target.width).toBe(last.bounds.width)
    expect(projectPrintDrawings(drawings, pages[1], geometry)).toEqual([])
  })

  it('reads saved types, axis ids, labels and caches, and updates the series through a session', async () => {
    const data = kept('chart-values')
    const [{ charts, issues }, { book }, { drawings }] = await Promise.all([
      readXlsxCharts(data), readXlsx(data), readXlsxDrawings(data),
    ])
    expect(issues).toEqual([])
    expect(charts).toHaveLength(1)
    const [chart] = charts
    expect(chart.frame).toMatchObject({ kind: 'chart', part: 'xl/charts/chart1.xml' })
    expect(drawings).toContainEqual(chart.frame)
    const [group] = chart.groups
    expect(group).toMatchObject({ type: 'bar', direction: 'col', grouping: 'clustered' })
    expect(group.axes).toEqual(chart.axes.map((axis) => axis.id))
    expect(group.series[0].categories?.points[0]).toEqual({ index: 0, value: 'East' })
    expect(group.series[0].values?.points[0]).toEqual({ index: 0, value: 200 })
    const session = createSession(book)
    session.setCells([{ sheet: 'Data', cell: 'B2', input: { value: 123 } }])
    const updated = resolveChartData(chart, session)
    expect(updated.groups[0].series[0].values?.points[0]).toEqual({ index: 0, value: 123 })
    expect(group.series[0].values?.points[0]).toEqual({ index: 0, value: 200 })
  })

  it('accepts a host resolver for a named source and retains its file cache when unresolved', async () => {
    const { charts: [chart] } = await readXlsxCharts(kept('chart-values'))
    const [series] = chart.groups[0].series
    series.values = { ...series.values!, reference: 'book.xlsx!DynamicValues' }
    const reader = { getValues: () => [['other']], resolve: (reference: string) => reference.endsWith('!DynamicValues') ? [7, null, 9] : undefined }
    const resolved = resolveChartData(chart, reader).groups[0].series[0]
    expect(resolved.values?.points).toEqual([{ index: 0, value: 7 }, { index: 1, value: null }, { index: 2, value: 9 }])
    expect(resolveChartData(chart, { getValues: reader.getValues }).groups[0].series[0].values).toBe(series.values)
  })

  it('can read a line chart from a dedicated chart sheet', async () => {
    const { charts, issues } = await readXlsxCharts(kept('chartsheet-values'))
    expect(issues).toEqual([])
    expect(charts[0].frame).toMatchObject({ sheet: 'Trend', sheetKind: 'chartsheet', anchor: { kind: 'absolute' } })
    expect(charts[0].groups[0].type).toBe('line')
  })
})
