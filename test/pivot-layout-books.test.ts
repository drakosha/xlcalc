/**
 * The layout against Excel's books (`test/excel/win/pivot/<set>/`): each PivotTable read, its calculation taken from what
 * Excel saved (`savedResult`), laid out again, gives Excel's saved lines and place, and the values and number formats
 * of the cells Excel showed (`<set>.json`, the area with the Filters area). A case of several phases has a book per
 * phase, `<case>-<n>`, the last one `<case>`. A PivotTable Excel 365 could not lay out over the cells it would take
 * shrinks to #SPILL! at its corner: written again (`placePivots`), it gives the same cells and definition.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { dateToSerial } from '../src/date-serial.js'
import { formatAddress, parseAddress } from '../src/model.js'
import { layoutPivot, type PivotCell } from '../src/pivot/layout.js'
import { placePivots } from '../src/pivot/place.js'
import { readXlsx } from '../src/xlsx.js'
import { savedResult } from './pivot-saved.js'

const ROOT = new URL('./excel/win/pivot/', import.meta.url).pathname
const SETS = ['basic', 'gpd-1', 'layout', 'filter', 'sort', 'edit', 'refresh', 'calc', 'layout-held', 'filter-held', 'sort-held', 'edit-held', 'show', 'show-items'].filter((s) => existsSync(`${ROOT}${s}.json`))

interface Shown {
  area: string
  values: unknown[][]
  formats: string[][]
}

/** A built-in number format as Excel's cell shows its code. */
const BUILTIN: Record<number, string> = { 4: '#,##0.00', 10: '0.00%', 14: 'm/d/yyyy' }

/** What a laid out cell holds as the answers write Excel's `Value2`: a date its serial, an error by its code, an empty cell null. */
function value2(c: PivotCell | undefined): unknown {
  if (!c) return null
  if (c.error) return { error: c.error }
  const v = c.value
  return typeof v === 'object' ? dateToSerial(v.date) : (v ?? null)
}

/** The answer of a case's book: the phase its name numbers, the last without a number. */
function answerOf(answers: Record<string, { phases: { pivots: Record<string, Shown> }[] }>, id: string) {
  const own = answers[id]
  if (own) return own.phases[own.phases.length - 1]
  const m = /^(.*)-(\d+)$/.exec(id)
  return m ? answers[m[1]]?.phases[Number(m[2]) - 1] : undefined
}

describe.each(SETS)('the layout of the PivotTables of set %s', (set) => {
  const { answers } = JSON.parse(readFileSync(`${ROOT}${set}.json`, 'utf8'))
  const books = readdirSync(`${ROOT}${set}`).filter((f) => f.endsWith('.xlsx'))
  it.each(books)('%s', async (file) => {
    const { book } = await readXlsx(readFileSync(`${ROOT}${set}/${file}`))
    const phase = answerOf(answers, file.replace(/\.xlsx$/, ''))
    let seen = 0
    for (const sheet of book.sheets)
      for (const t of sheet.pivotTables ?? []) {
        const shown = phase?.pivots[t.name] ?? phase?.pivots[`${sheet.name}!${t.name}`]
        if (!shown) continue
        seen++
        const cache = book.pivotCaches!.find((c) => c.id === t.cache)!
        // the saved cells hold the values as shown, Show Values As applied
        const saved = savedResult(t, cache, sheet)
        const laid = layoutPivot(saved, book, saved.value)
        expect(laid.rowLines).toEqual(t.rowLines)
        expect(laid.columnLines).toEqual(t.columnLines)
        if (sheet.cells[t.location.range]?.formula === '#SPILL!') {
          // a refresh Excel 365 could not lay out over the cells there: the PivotTable shrunk to #SPILL! at its corner
          const r = placePivots(book, [{ sheet: sheet.name, table: t, laid }])
          if (!r.ok) throw new Error(r.message)
          const after = r.book.sheets.find((x) => x.name === sheet.name)!
          expect(after.pivotTables!.find((x) => x.name === t.name)).toEqual(t)
          expect(after.cells).toEqual(sheet.cells)
          continue
        }
        expect(laid.location).toEqual(t.location)
        const [a, b] = shown.area.split(':').map(parseAddress)
        // the blank row under the Filters area holds what was put there
        const blank = t.location.pageRows ? parseAddress(t.location.range.split(':')[0]).row - 1 : -1
        const mine = new Map(laid.cells.map((c) => [formatAddress(c.row, c.col), c]))
        const got: unknown[][] = []
        const formats: string[][] = []
        for (let r = a.row; r <= (b ?? a).row; r++) {
          got.push([])
          formats.push([])
          for (let c = a.col; c <= (b ?? a).col; c++) {
            const cell = mine.get(formatAddress(r, c))
            got[got.length - 1].push(r === blank ? shown.values[r - a.row][c - a.col] : value2(cell))
            const f = cell?.numberFormat
            formats[formats.length - 1].push(f === undefined ? 'General' : typeof f === 'number' ? BUILTIN[f] : f)
          }
        }
        const want = shown.values.map((row, i) => row.map((v, j) => (typeof v === 'number' && typeof got[i][j] === 'number' && Math.abs(v - (got[i][j] as number)) < 1e-9 ? got[i][j] : v)))
        expect(got).toEqual(want)
        // a format as the file stores it escapes its literal characters (`yyyy\-mm\-dd`), as COM gives it not
        const unescaped = (g: string[][]) => g.map((r) => r.map((f) => f.replace(/\\(.)/g, '$1')))
        expect(unescaped(formats)).toEqual(unescaped(shown.formats))
        expect(laid.cells.every((c) => c.row >= a.row && c.row <= (b ?? a).row && c.col >= a.col && c.col <= (b ?? a).col)).toBe(true)
      }
    expect(seen).toBeGreaterThan(0)
  })
})
