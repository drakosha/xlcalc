/**
 * The engine's prediction of an alerts set (`alerts`, `alerts-held`), written before Excel's answers are read: each
 * case's pivot table laid out by the engine alone (`drive.ts`) up to its first `read`, then its one change made in a
 * session with `xlcalc/pivot` loaded as a host makes it: an operation of the structure through `editStructure`, text
 * typed into several cells through `fillCellText` (Ctrl+Enter), set through `setCells`, cleared through `clearCells`,
 * pasted through `pasteText`. Each case gives the refusal's code, `reason` and message, the error an edit threw with its
 * `reason`, or the cells of every pivot table's area after it (keyed by name, as the harness reads them). Run: `npx tsx
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
  | { pivots: Record<string, CellValue[][]> }

/** The operation of the structure a `struct` operation of the harness is, on `P` unless it names another sheet. */
function structureOp(o: Extract<PivotOp, { struct: string }>): StructureOp {
  const sheet = o.sheet ?? 'P'
  const range = o.range ?? ''
  const row = Number(/\d+/.exec(range)?.[0])
  const column = /^[A-Z]+/.exec(range)?.[0] ?? 'A'
  switch (o.struct) {
    case 'insertRows':
    case 'deleteRows':
      return { op: o.struct, sheet, row }
    case 'insertColumns':
    case 'deleteColumns':
      return { op: o.struct, sheet, column }
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

/** The case's one change made in the session, as a host makes it. */
async function changed(s: Session, o: PivotOp): Promise<AlertPrediction> {
  try {
    type Answer = { ok: boolean; code?: string; reason?: string; message?: string }
    let r: Answer
    if ('struct' in o) r = s.editStructure(structureOp(o))
    else if ('paste' in o) r = s.pasteText('P', o.at, o.paste.replace(/\n/g, '\r\n'))
    else if ('edit' in o) {
      const [sheet, range] = o.edit.split('!')
      const to = o.to
      if (to === null) r = { ok: true, ...s.clearCells([{ sheet, range }], 'contents') }
      else if (typeof to === 'object' && 'typed' in to) r = s.fillCellText(sheet, range, to.typed)
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
  return { pivots: Object.fromEntries((sheet.pivotTables ?? []).map((t) => [t.name, shown(book, 'P', rectText(pivotArea(t)!))])) }
}

/** The addresses of a range's cells, by rows. */
function rangeCells(range: string): string[] {
  const x = parseRange(range)!
  return Array.from({ length: (x.r2 - x.r1 + 1) * (x.c2 - x.c1 + 1) }, (_, k) => formatAddress(x.r1 + Math.floor(k / (x.c2 - x.c1 + 1)), x.c1 + (k % (x.c2 - x.c1 + 1))))
}

if (import.meta.url === `file://${process.argv[1]}`) {
  registerPivot(pivotEngine)
  const set = process.argv[2]
  const { buildPivotCases } = (await import(`./${set}.ts`)) as PivotSet
  const out: Record<string, AlertPrediction> = {}
  for (const c of buildPivotCases() as PivotCase[]) {
    const k = c.ops.findIndex((o) => 'read' in o)
    const book = drive({ ...c, ops: c.ops.slice(0, k) }).at(-1)!.book
    out[c.id] = await changed(createSession(book, { clock: { now: '2026-10-31T12:00:00Z', seed: 0 } }), c.ops[k + 1])
  }
  writeFileSync(new URL(`${set}.predicted.json`, import.meta.url), JSON.stringify(out, null, 1) + '\n')
  console.log(`${Object.keys(out).length} cases, ${Object.values(out).filter((v) => !('pivots' in v)).length} refused`)
}
