/**
 * A pivot set's cases replayed on the engine, phase by phase, over the books Excel saved (`test/excel/win/pivot/<set>/`,
 * a book per phase: `<case>-1.xlsx`…, the last `<case>.xlsx`): each phase's PivotTables laid out again from the cache
 * Excel saved, the first phase's refreshed from its source too, and a phase that refreshes after edits of the source made
 * from the phase before. The test against Excel's answers (`test/pivot-excel.test.ts`) and the engine's predictions
 * (`predict.ts`) both go through it.
 */
import { existsSync, readFileSync } from 'node:fs'
import type { BookInput, CellValue, PivotRef } from '../../../src/contract.js'
import { createState } from '../../../src/engine.js'
import { dateToSerial, formatAddress } from '../../../src/model.js'
import { refreshPivots } from '../../../src/pivot/engine.js'
import { clockOptions } from '../../../src/result.js'
import { parseRange } from '../../../src/structure/area.js'
import type { Applied } from '../../../src/structure/geometry.js'
import { stateValues } from '../../../src/structure/values.js'
import { ERROR_CODE } from '../../../src/syntax/errors.js'
import { readXlsx } from '../../../src/xlsx.js'
import type { PivotCase, PivotOp } from './cases.js'

const DIR = new URL('../../../test/excel/win/pivot/', import.meta.url)
/** The day the sets ran in Excel, for relative date filters. */
const CLOCK = { now: '2026-10-30', seed: 0 }

/** The values of `range` on `sheet` as its cells hold them. */
export function shown(book: BookInput, sheet: string, range: string): CellValue[][] {
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

/** Numbers to 15 significant digits, as the answers compare. */
export const rounded = (g: CellValue[][]) => g.map((r) => r.map((v) => (typeof v === 'number' ? Number(v.toPrecision(15)) : v)))

/** A case's operations by phase: each `read` ends one. */
export function phases(c: PivotCase): PivotOp[][] {
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
async function book(set: string, id: string, k: number, last: number): Promise<BookInput | null> {
  const file = new URL(`${set}/${id}${k < last ? `-${k + 1}` : ''}.xlsx`, DIR)
  return existsSync(file) ? (await readXlsx(readFileSync(file))).book : null
}

/** A book's PivotTables refreshed (`fromSource`) or laid out again. */
function laid(b: BookInput, refs: PivotRef[] | undefined, fromSource: boolean): Applied {
  const values = stateValues(createState(b, clockOptions(CLOCK)))
  return refreshPivots(b, refs, values, values.today!(), fromSource)
}

/** One way a phase is made: its name (`<case>/<phase>/<how>`), the PivotTables laid out, and whether Excel's own refusals count. */
export interface Replay {
  at: string
  result: Applied
  /** A refresh after edits, whose refusal Excel may share; a book laid out again Excel never refused. */
  refresh: boolean
}

/** The replays of a case's phases, for `count` phases (the answer's, or the books there are). */
export async function replays(set: string, c: PivotCase, count: number): Promise<Replay[]> {
  const ops = phases(c)
  const out: Replay[] = []
  for (let k = 0; k < count; k++) {
    const b = await book(set, c.id, k, count - 1)
    if (!b) continue
    out.push({ at: `${c.id}/${k}/saved`, result: laid(b, undefined, false), refresh: false })
    if (k === 0 && !c.from && !c.file && fresh(ops[0])) out.push({ at: `${c.id}/0/source`, result: laid(b, undefined, true), refresh: false })
    const refs = k > 0 ? refreshed(ops[k]) : null
    const prev = refs && (await book(set, c.id, k - 1, count - 1))
    if (!prev) continue
    // the PivotTables and caches of the phase before, over the source as this phase's book holds it, with the phase's
    // edits of the PivotTables' sheet (a cell in the way cleared)
    const pivots = prev.sheets.find((x) => x.name === 'P')!
    const cells = { ...pivots.cells }
    for (const op of ops[k]) if ('edit' in op && op.edit.startsWith('P!')) op.to === null ? delete cells[op.edit.slice(2)] : (cells[op.edit.slice(2)] = typeof op.to === 'object' ? ('formula' in op.to ? { formula: op.to.formula } : { value: op.to.typed }) : { value: op.to })
    const source = { ...b, sheets: b.sheets.map((s) => (s.name === 'P' ? { ...pivots, cells } : s)), pivotCaches: prev.pivotCaches }
    out.push({ at: `${c.id}/${k}/refresh`, result: laid(source, refs === 'all' ? undefined : refs!, true), refresh: true })
  }
  return out
}

/** The number of phases of a case whose books are saved: one per `read`, and the last. */
export const phaseCount = (c: PivotCase) => phases(c).length
