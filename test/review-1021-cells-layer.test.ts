/**
 * Reproductions from the 21.10 review, cell looks read row by row (`CellsLayer`): a cell given twice keeps the last look
 * given it, also in a row given again after later rows, and so a file read with such a row does.
 */
import { describe, expect, it } from 'vitest'
import { CellsLayer } from '../src/looks/layer.js'
import { cellsOf } from './looks-book.js'
import { sheet } from './xlsx-fixtures.js'

/** The styles of cells given as [row, column, style] in turn, read back from the layer at each cell. */
function lastLooks(seq: [number, number, number][]): [string, number | undefined][] {
  const cl = new CellsLayer()
  for (const [r, c, style] of seq) cl.add(r, c, cl.looks.id(style, undefined))
  const l = cl.build()
  const cells = [...new Set(seq.map(([r, c]) => `${r},${c}`))]
  return cells.map((rc) => {
    const [r, c] = rc.split(',').map(Number)
    return [rc, l.lookAt(r, c)?.style]
  })
}

/** What the sequence should give: each cell's last style. */
const expected = (seq: [number, number, number][]) => [...new Map(seq.map(([r, c, s]) => [`${r},${c}`, s]))]

describe('CellsLayer: the last look given a cell counts', () => {
  it.each([
    ['in order', [[1, 0, 1], [1, 0, 2]]],
    ['a row given again, a cell twice within the visit', [[1, 0, 1], [2, 0, 1], [1, 0, 1], [1, 0, 2]]],
    ['a row given again, only its cell twice', [[2, 0, 1], [1, 3, 1], [1, 3, 2]]],
    ['a row given again, two cells, one twice', [[2, 0, 1], [1, 3, 2], [1, 4, 2], [1, 3, 1]]],
    ['a row given again twice', [[1, 0, 1], [2, 0, 1], [1, 0, 2], [3, 0, 1], [1, 0, 1], [1, 0, 2]]],
  ] as [string, [number, number, number][]][])('%s', (_, seq) => {
    expect(lastLooks(seq)).toEqual(expected(seq))
  })

  it('a file with a row given again after a later one: the cell takes the last look given it', async () => {
    const styles = '<cellXfs count="3"><xf numFmtId="0"/><xf numFmtId="2"/><xf numFmtId="10"/></cellXfs>'
    const r = await sheet('<row r="3"><c r="A3" s="1"/></row><row r="2"><c r="A2" s="1"/><c r="A2" s="2"/></row>', { styles })
    expect(cellsOf(r.book.sheets[0]).A2).toEqual({ numberFormat: 10 })
  })
})
