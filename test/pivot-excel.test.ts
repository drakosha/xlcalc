/**
 * PivotTables against Excel for Windows: each book of a set of `bench/coverage/pivot` read, its PivotTable laid out
 * again by `xlcalc/pivot` from the cache Excel saved and refreshed from its source, and the cells of its area against
 * those Excel showed (`test/excel/win/pivot/<set>.json`), but the cases each set lists as known differences.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { BookInput, CellValue } from '../src/contract.js'
import { createState } from '../src/engine.js'
import { dateToSerial, formatAddress } from '../src/model.js'
import { ERROR_CODE } from '../src/syntax/errors.js'
import { refreshPivots } from '../src/pivot/engine.js'
import { clockOptions } from '../src/result.js'
import { parseRange } from '../src/structure/area.js'
import { stateValues } from '../src/structure/values.js'
import { readXlsx } from '../src/xlsx.js'

const DIR = new URL('./excel/win/pivot/', import.meta.url)

interface Area {
  area: string
  values: CellValue[][]
}

/**
 * The cases of a set whose PivotTable differs from Excel's, by set. basic: the labels of error items (`#N/A`), which
 * Excel writes as text; empty cells with "For empty cells show", which Excel shows as 0 (both of the layout).
 */
const KNOWN: Record<string, string[]> = { basic: ['amount-rows', 'shown-errors'] }

/** The values of `range` on the sheet `P` of a book as its cells hold them. */
function shown(book: BookInput, range: string): CellValue[][] {
  const x = parseRange(range)!
  const cells = book.sheets.find((s) => s.name === 'P')!.cells
  return Array.from({ length: x.r2 - x.r1 + 1 }, (_, r) =>
    Array.from({ length: x.c2 - x.c1 + 1 }, (_, c) => {
      const cell = cells[formatAddress(x.r1 + r, x.c1 + c)]
      const v = cell?.formula?.startsWith("#") ? { error: ERROR_CODE[cell.formula] } : cell?.value
      return v === undefined ? null : typeof v === 'object' && v !== null && 'date' in v ? dateToSerial(v.date) : (v as CellValue)
    }),
  )
}

/** Numbers to 15 significant digits, as Excel shows them. */
const rounded = (g: CellValue[][]) => g.map((r) => r.map((v) => (typeof v === 'number' ? Number(v.toPrecision(15)) : v)))

async function differences(set: string, fromSource: boolean): Promise<string[]> {
  const answers = JSON.parse(readFileSync(new URL(`${set}.json`, DIR), 'utf8')) as { answers: Record<string, { phases: { pivots: Record<string, Area | { error: string }> }[] }> }
  const out: string[] = []
  for (const [id, answer] of Object.entries(answers.answers)) {
    const want = answer.phases[0]?.pivots.PT1
    if (!want || 'error' in want) continue
    const { book } = await readXlsx(readFileSync(new URL(`${set}/${id}.xlsx`, DIR)))
    const values = stateValues(createState(book, clockOptions({ now: '2026-10-30', seed: 0 })))
    const r = refreshPivots(book, [{ sheet: 'P', name: 'PT1' }], values, values.today!(), fromSource)
    if (!r.ok) {
      out.push(`${id}: ${r.message}`)
      continue
    }
    const got = rounded(shown(r.book, want.area))
    if (JSON.stringify(got) !== JSON.stringify(rounded(want.values))) out.push(`${id}: ${JSON.stringify(got)} for ${JSON.stringify(want.values)}`)
  }
  return out.sort()
}

describe('PivotTables against Excel for Windows', () => {
  for (const set of Object.keys(KNOWN))
    for (const fromSource of [false, true])
      it(`${set}, ${fromSource ? 'refreshed from the source' : 'laid out from the saved cache'}`, async () => {
        const known = KNOWN[set]
        const found = await differences(set, fromSource)
        expect(found.filter((d) => !known.some((k) => d.startsWith(`${k}:`)))).toEqual([])
      }, 120_000)
})
