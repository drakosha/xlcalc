/**
 * The engine's prediction of an alerts set (`alerts`, `alerts-clear`, `alerts-held`), written before Excel's answers are
 * read: each case's pivot tables laid out by the engine alone (`drive.ts`) up to its first `read`, the values it puts
 * beside them on `P` set in a session with `xlcalc/pivot` loaded, then its one change made there as a host makes it: an
 * operation of the structure through `editStructure`, text typed or a value set into several cells through
 * `fillCellText` (Ctrl+Enter), into one through `setCells`, a range cleared (its contents, everything or its formats)
 * through `clearCells`, text pasted through `pasteText`, a copy pasted through `copyRange`. Each case gives the
 * refusal's code, `reason` and message, the error an edit threw with its `reason`, or the cells of every pivot table's
 * area after it (keyed by name, as the harness reads them) and the cells it watches. Run: `npx tsx
 * bench/coverage/pivot/alerts-predict.ts <set>`; it writes `bench/coverage/pivot/<set>.predicted.json`. The comparison
 * with Excel: `test/pivot-alerts.test.ts`.
 */
import { writeFileSync } from 'node:fs'
import { createSession, PivotInputError, registerPivot, type CellValue, type Session, type StructureOp } from '../../../src/index.js'
import { pivotArea } from '../../../src/pivot/area.js'
import { pivotEngine } from '../../../src/pivot/index.js'
import { formatAddress } from '../../../src/model.js'
import { parseRange, rectText } from '../../../src/structure/area.js'
import type { PivotCase, PivotOp, PivotSet } from './cases.js'
import { drive } from './drive.js'
import { shown } from './phases.js'

/** What the engine does with a case's change. */
export type AlertPrediction =
  | { refused: { code: string; reason?: string; message: string } }
  | { threw: { name: string; reason?: string; message: string } }
  | { pivots: Record<string, CellValue[][]>; cells?: Record<string, CellValue> }

/** The operation of the structure a `struct` operation of the harness is, on `P` unless it names another sheet. */
function structureOp(o: Extract<PivotOp, { struct: string }>): StructureOp {
  const sheet = o.sheet ?? 'P'
  const range = o.range ?? ''
  // the entire rows or columns of the range, as Range.EntireRow and EntireColumn take them
  const x = parseRange(range)
  switch (o.struct) {
    case 'insertRows':
    case 'deleteRows':
      return { op: o.struct, sheet, row: x!.r1 + 1, count: x!.r2 - x!.r1 + 1 }
    case 'insertColumns':
    case 'deleteColumns':
      return { op: o.struct, sheet, column: formatAddress(0, x!.c1).replace(/\d+$/, ''), count: x!.c2 - x!.c1 + 1 }
    case 'copyRange':
      return { op: 'copyRange', from: { sheet, range }, to: { sheet: o.target ?? 'P', range: o.to! } }
    case 'moveRange':
      return { op: 'moveRange', from: { sheet, range }, to: { sheet: o.target ?? 'P', cell: o.to! } }
    case 'fillDown':
      return { op: 'fillRange', sheet: o.target ?? 'P', range: o.to!.split(':')[0], to: o.to! }
    case 'mergeCells':
      return { op: 'mergeCells', sheet, range }
    case 'addTable': {
      const x = parseRange(range)!
      return { op: 'addTable', sheet, table: { name: o.target!, range, columns: Array.from({ length: x.c2 - x.c1 + 1 }, (_, k) => ({ name: `Column${k + 1}` })) } }
    }
  }
}

/** The ranges of a range of areas apart by commas (`A1:B1,A3:D10`). */
const areasOf = (sheet: string, range: string) => range.split(',').map((x) => ({ sheet, range: x }))

/** The case's one change made in the session, as a host makes it; with the cells it watches after it. */
async function changed(s: Session, o: PivotOp, watch: string[] = []): Promise<AlertPrediction> {
  try {
    type Answer = { ok: boolean; code?: string; reason?: string; message?: string }
    let r: Answer
    if ('struct' in o) r = s.editStructure(structureOp(o))
    else if ('paste' in o) r = s.pasteText('P', o.at, o.paste.replace(/\n/g, '\r\n'))
    else if ('copyPaste' in o) {
      const [sheet, range] = o.copyPaste.split('!')
      r = s.editStructure({ op: 'copyRange', from: { sheet, range }, to: { sheet: 'P', range: o.at } })
    } else if ('clear' in o) {
      const [sheet, range] = o.clear.split('!')
      r = { ok: true, ...s.clearCells(areasOf(sheet, range), o.what) }
    } else if ('edit' in o) {
      const [sheet, range] = o.edit.split('!')
      const to = o.to
      if (to === null) r = { ok: true, ...s.clearCells(areasOf(sheet, range), 'contents') }
      else if (typeof to === 'object' && 'typed' in to) r = s.fillCellText(sheet, range, to.typed)
      // a value set over a range at once is an entry into a selection too (the macro recorder's Ctrl+Enter)
      else if (range.includes(':') && typeof to !== 'object') r = s.fillCellText(sheet, range, String(to))
      else {
        r = { ok: true, ...s.setCells(rangeCells(range).map((cell) => ({ sheet, cell, input: typeof to === 'object' ? { formula: to.formula } : { value: to } }))) }
      }
    } else throw new Error(`no change in ${JSON.stringify(o)}`)
    if (!r.ok) return { refused: { code: r.code!, ...(r.reason ? { reason: r.reason } : {}), message: r.message! } }
  } catch (e) {
    const err = e as Error
    return { threw: { name: err.name, ...(e instanceof PivotInputError ? { reason: e.reason } : {}), message: err.message } }
  }
  const book = s.getBook()
  const sheet = book.sheets.find((x) => x.name === 'P')!
  const pivots = Object.fromEntries((sheet.pivotTables ?? []).map((t) => [t.name, shown(book, 'P', rectText(pivotArea(t)!))]))
  if (!watch.length) return { pivots }
  const cells = Object.fromEntries(watch.flatMap((w) => {
    const [name, range] = w.split('!')
    const x = parseRange(range)!
    return shown(book, name, range).flatMap((row, r) => row.map((v, c) => [`${name}!${formatAddress(x.r1 + r, x.c1 + c)}`, v]))
  }))
  return { pivots, cells }
}

/** The addresses of a range's cells, by rows. */
function rangeCells(range: string): string[] {
  const x = parseRange(range)!
  return Array.from({ length: (x.r2 - x.r1 + 1) * (x.c2 - x.c1 + 1) }, (_, k) => formatAddress(x.r1 + Math.floor(k / (x.c2 - x.c1 + 1)), x.c1 + (k % (x.c2 - x.c1 + 1))))
}

/** A case's change predicted: its pivot tables played up to its first `read`, then the operation after it in a session. */
export async function predictedAlert(c: PivotCase): Promise<AlertPrediction> {
  registerPivot(pivotEngine)
  const k = c.ops.findIndex((o) => 'read' in o)
  // the pivot tables played on the engine; values put on `P` beside them, set in the session
  const onP = (o: PivotOp) => 'edit' in o && o.edit.startsWith('P!')
  const book = drive({ ...c, ops: c.ops.slice(0, k).filter((o) => !onP(o)) }).at(-1)!.book
  const s = createSession(book, { clock: { now: '2026-10-31T12:00:00Z', seed: 0 } })
  for (const o of c.ops.slice(0, k).filter(onP) as Extract<PivotOp, { edit: string }>[]) s.setCells([{ sheet: 'P', cell: o.edit.slice(2), input: { value: o.to as string | number } }])
  return changed(s, c.ops[k + 1], c.watch)
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const set = process.argv[2]
  const { buildPivotCases } = (await import(`./${set}.ts`)) as PivotSet
  const out: Record<string, AlertPrediction> = {}
  for (const c of buildPivotCases() as PivotCase[]) out[c.id] = await predictedAlert(c)
  writeFileSync(new URL(`${set}.predicted.json`, import.meta.url), JSON.stringify(out, null, 1) + '\n')
  console.log(`${Object.keys(out).length} cases, ${Object.values(out).filter((v) => !('pivots' in v)).length} refused`)
}
