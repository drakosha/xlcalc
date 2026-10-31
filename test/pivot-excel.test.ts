/**
 * PivotTables against Excel for Windows, over the books and answers of the sets of `bench/coverage/pivot`
 * (`test/excel/win/pivot/<set>.json`, a book per phase: `<case>-1.xlsx`…, the last `<case>.xlsx`). Each phase's book is read and
 * its PivotTables laid out again by `xlcalc/pivot` from the cache Excel saved; the first phase's refreshed from its
 * source too; and a phase that refreshes after edits of the source is made from the phase before: its PivotTables and
 * caches, the source as this phase's book holds it, refreshed. The cells of each PivotTable's area are compared with
 * those Excel showed, numbers to 15 significant digits, a refusal with Excel's; the cases each set lists as known
 * differences are left out.
 */
import { existsSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import * as basic from '../bench/coverage/pivot/basic.js'
import type { PivotCase, PivotOp, PivotPhase, PivotSet } from '../bench/coverage/pivot/cases.js'
import * as filter from '../bench/coverage/pivot/filter.js'
import * as filterHeld from '../bench/coverage/pivot/filter-held.js'
import * as layout from '../bench/coverage/pivot/layout.js'
import * as layoutHeld from '../bench/coverage/pivot/layout-held.js'
import * as refresh from '../bench/coverage/pivot/refresh.js'
import * as show from '../bench/coverage/pivot/show.js'
import * as showHeld from '../bench/coverage/pivot/show-held.js'
import * as showItems from '../bench/coverage/pivot/show-items.js'
import * as sort from '../bench/coverage/pivot/sort.js'
import * as sortHeld from '../bench/coverage/pivot/sort-held.js'
import type { BookInput, CellValue, PivotRef } from '../src/contract.js'
import { createState } from '../src/engine.js'
import { dateToSerial, formatAddress } from '../src/model.js'
import { refreshPivots } from '../src/pivot/engine.js'
import { clockOptions } from '../src/result.js'
import { parseRange } from '../src/structure/area.js'
import { stateValues } from '../src/structure/values.js'
import { ERROR_CODE } from '../src/syntax/errors.js'
import { readXlsx } from '../src/xlsx.js'

const DIR = new URL('./excel/win/pivot/', import.meta.url)
/** The day the sets ran in Excel, for relative date filters. */
const CLOCK = { now: '2026-10-30', seed: 0 }

/**
 * The cases of a set whose PivotTables differ from Excel's, by set, as `<case>/<phase>/<how>`. sort: items of several
 * custom lists at once, which Excel orders by no order we found (`docs/development.md`).
 */
const KNOWN: Record<string, string[]> = {
  'layout-held': [],
  'filter-held': [],
  'sort-held': [],
  basic: [],
  filter: [],
  sort: ['months-down/0/saved', 'months-down/0/source'],
  layout: [],
  refresh: [],
  show: [],
  'show-held': [],
  'show-items': [],
}

const SETS: Record<string, PivotSet> = { basic, filter, sort, refresh, layout, 'layout-held': layoutHeld, 'filter-held': filterHeld, 'sort-held': sortHeld, show, 'show-items': showItems, 'show-held': showHeld }

/** The values of `range` on `sheet` as its cells hold them. */
function shown(book: BookInput, sheet: string, range: string): CellValue[][] {
  const x = parseRange(range)!
  const cells = book.sheets.find((s) => s.name === sheet)!.cells
  return Array.from({ length: x.r2 - x.r1 + 1 }, (_, r) =>
    Array.from({ length: x.c2 - x.c1 + 1 }, (_, c) => {
      const cell = cells[formatAddress(x.r1 + r, x.c1 + c)]
      if (cell?.formula?.startsWith('#')) return { error: ERROR_CODE[cell.formula] }
      const v = cell?.value
      return v === undefined ? null : typeof v === 'object' && v !== null && 'date' in v ? dateToSerial(v.date) : (v as CellValue)
    }),
  )
}

const rounded = (g: CellValue[][]) => g.map((r) => r.map((v) => (typeof v === 'number' ? Number(v.toPrecision(15)) : v)))

/**
 * The relative difference a set's numbers may have beyond their 15 significant digits: Show Values As divides by sums
 * whose last bit Excel's order of adding sets (`docs/limitations.md`), which a share carries into its 15th digit.
 */
const TOLERANCE: Record<string, number> = { show: 1e-14, 'show-items': 1e-14, 'show-held': 1e-14 }

/** The cells alike: the same values, numbers to 15 significant digits or within `tolerance` of each other. */
const alike = (got: CellValue[][], want: CellValue[][], tolerance: number) =>
  got.length === want.length &&
  got.every((r, i) => r.length === want[i].length && r.every((v, j) => {
    const w = want[i][j]
    return JSON.stringify(v) === JSON.stringify(w) || (typeof v === 'number' && typeof w === 'number' && Math.abs(v - w) <= tolerance * Math.abs(w))
  }))

/** A case's operations by phase: each `read` ends one. */
function phases(c: PivotCase): PivotOp[][] {
  const out: PivotOp[][] = [[]]
  for (const op of c.ops) {
    if ('read' in op) out.push([])
    else out[out.length - 1].push(op)
  }
  return out
}

/** The PivotTables a phase refreshes from the source after edits alone (`all` for Refresh All), or null for a phase that does more. */
function refreshed(ops: PivotOp[]): PivotRef[] | 'all' | null {
  let refs: PivotRef[] | 'all' | null = null
  for (const op of ops) {
    if ('refreshAll' in op) refs = 'all'
    else if ('refresh' in op) refs = [{ sheet: 'P', name: op.pt ?? 'PT1' }]
    else if (!('edit' in op || 'resize' in op || 'define' in op)) return null
  }
  return refs
}

/** Whether the source is as the cache last read it: no edit after the last refresh or field placed. */
const fresh = (ops: PivotOp[]) => ops.findLastIndex((o) => 'edit' in o) < ops.findLastIndex((o) => 'refresh' in o || 'refreshAll' in o || 'rows' in o || 'columns' in o || 'value' in o)

/** The book saved at the end of phase `k` of `last`: a copy `<case>-<k + 1>.xlsx`, the last one the case's own. */
const book = async (set: string, id: string, k: number, last: number) => {
  const file = new URL(`${set}/${id}${k < last ? `-${k + 1}` : ''}.xlsx`, DIR)
  return existsSync(file) ? (await readXlsx(readFileSync(file))).book : null
}

/** A book's PivotTables refreshed or laid out again, each against the phase's answer. */
function compared(at: string, b: BookInput, refs: PivotRef[] | undefined, fromSource: boolean, want: PivotPhase, tolerance: number): string[] {
  const values = stateValues(createState(b, clockOptions(CLOCK)))
  const r = refreshPivots(b, refs, values, values.today!(), fromSource)
  const excelRefused = (want.errors ?? []).length > 0
  if (!r.ok) return excelRefused ? [] : [`${at}: refused, ${r.message}`]
  if (excelRefused) return [`${at}: Excel refused (${want.errors!.join('; ')})`]
  return Object.entries(want.pivots).flatMap(([name, area]) => {
    if ('error' in area) return []
    const got = rounded(shown(r.book, 'P', area.area))
    return alike(got, rounded(area.values), tolerance) ? [] : [`${at} ${name}: ${JSON.stringify(got)} for ${JSON.stringify(area.values)}`]
  })
}

async function differences(set: string): Promise<string[]> {
  const answers = JSON.parse(readFileSync(new URL(`${set}.json`, DIR), 'utf8')) as { answers: Record<string, { phases: PivotPhase[] }> }
  const cases = SETS[set].buildPivotCases()
  const tolerance = TOLERANCE[set] ?? 0
  const out: string[] = []
  for (const c of cases) {
    const answer = answers.answers[c.id]
    if (!answer) continue
    const ops = phases(c)
    for (const [k, want] of answer.phases.entries()) {
      const last = answer.phases.length - 1
      const b = await book(set, c.id, k, last)
      if (!b) continue
      out.push(...compared(`${c.id}/${k}/saved`, b, undefined, false, { ...want, errors: [] }, tolerance))
      if (k === 0 && !c.from && !c.file && fresh(ops[0])) out.push(...compared(`${c.id}/0/source`, b, undefined, true, { ...want, errors: [] }, tolerance))
      const refs = k > 0 ? refreshed(ops[k]) : null
      const prev = refs && (await book(set, c.id, k - 1, last))
      if (!prev) continue
      // the PivotTables and caches of the phase before, over the source as this phase's book holds it, with the phase's
      // edits of the PivotTables' sheet (a cell in the way cleared)
      const pivots = prev.sheets.find((x) => x.name === 'P')!
      const cells = { ...pivots.cells }
      for (const op of ops[k]) if ('edit' in op && op.edit.startsWith('P!')) op.to === null ? delete cells[op.edit.slice(2)] : (cells[op.edit.slice(2)] = typeof op.to === 'object' ? ('formula' in op.to ? { formula: op.to.formula } : { value: op.to.typed }) : { value: op.to })
      const source = { ...b, sheets: b.sheets.map((s) => (s.name === 'P' ? { ...pivots, cells } : s)), pivotCaches: prev.pivotCaches }
      out.push(...compared(`${c.id}/${k}/refresh`, source, refs === 'all' ? undefined : refs!, true, want, tolerance))
    }
  }
  return out.sort()
}

describe('PivotTables against Excel for Windows', () => {
  for (const set of Object.keys(KNOWN))
    it(set, async () => {
      const found = await differences(set)
      expect(found.filter((d) => !KNOWN[set].some((k) => d.startsWith(`${k}:`) || d.startsWith(`${k} `)))).toEqual([])
    }, 120_000)
})
