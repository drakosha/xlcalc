import { describe, expect, it } from 'vitest'
import { layoutPrintPages, projectPrintRect } from '../src/xlsx-print.js'

const margins = { left: 0, right: 0, top: 0, bottom: 0, header: 0, footer: 0 }

describe('worksheet coordinates on a print page', () => {
  it('maps within each measured line when the printer alters row and column pitches', () => {
    const row = [0, 10, 30]
    const column = [0, 20, 40]
    const geometry = {
      usedRange: 'A1:B2',
      rows: { count: 2, edge: (line: number) => row[line - 1] },
      columns: { count: 2, edge: (line: number) => column[line - 1] },
    }
    const deviceGeometry = {
      rowEdge: (line: number, zoom: number) => [0, 8, 40][line - 1] * zoom,
      columnEdge: (line: number, zoom: number) => [0, 24, 40][line - 1] * zoom,
    }
    const { pages } = layoutPrintPages({ sheet: 'S', margins }, geometry, { width: 60, height: 60, deviceGeometry })
    expect(projectPrintRect({ x: 10, y: 5, width: 20, height: 15 }, pages[0], geometry, deviceGeometry)).toEqual([
      { source: { x: 10, y: 5, width: 20, height: 15 }, target: { x: 12, y: 4, width: 20, height: 20 } },
    ])
  })

  it('projects each visible fragment where pages repeat a title and skip a hidden row', () => {
    const edges = [0, 10, 20, 20, 30, 40]
    const geometry = {
      usedRange: 'A1:A5',
      rows: { count: 5, edge: (line: number) => edges[line - 1] },
      columns: { count: 1, edge: (line: number) => (line - 1) * 40 },
    }
    const settings = { sheet: 'S', margins, repeatRows: { first: 1, last: 1 } }
    const { pages } = layoutPrintPages(settings, geometry, { width: 45, height: 21 })
    expect(pages[1].regions.map((region) => region.range)).toEqual(['A1', 'A4'])
    expect(projectPrintRect({ x: 30, y: 5, width: 25, height: 20 }, pages[1], geometry)).toEqual([
      { source: { x: 30, y: 5, width: 10, height: 5 }, target: { x: 30, y: 5, width: 10, height: 5 } },
      { source: { x: 30, y: 20, width: 10, height: 5 }, target: { x: 30, y: 10, width: 10, height: 5 } },
    ])
  })
})
