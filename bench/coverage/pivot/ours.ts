/**
 * The engine's files of the books of a pivot set, for the sets that open them in Excel (`write-1`, a case's `file`):
 * each book Excel saved for a case of the set (`test/excel/win/pivot/<set>/<id>.xlsx`) read and written again, by
 * `writeXlsx` (`xlsx`) or by `openKept` and `writeKept` as it is (`kept`); or, for a set whose cases edit a cell after
 * their first phase (`edit`), the book of that phase (`<id>-1.xlsx`) opened by `openKept`, the cell edited in a session
 * with `xlcalc/pivot` as the case edits it, and written by `writeKept` (`kept-edit`); into
 * `bench/windows/out/pivot-ours/<how>/<set>/`.
 *
 *   npx tsx bench/coverage/pivot/ours.ts basic xlsx
 *   npx tsx bench/coverage/pivot/ours.ts edit kept-edit
 */
import { mkdirSync, readFileSync, readdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import type { EditValue, PivotCase } from './cases.js'
import { createSession, registerPivot } from '../../../src/index.js'
import { pivotEngine } from '../../../src/pivot/index.js'
import { openKept, writeKept } from '../../../src/xlsx-keep.js'
import { readXlsx, writeXlsx } from '../../../src/xlsx.js'
import { setModule } from '../stored.js'
import type { PivotSet } from './cases.js'

export type Rewrite = 'xlsx' | 'kept' | 'kept-edit'

/** Where the engine's file of a case's book is, from the package's folder. */
export const oursPath = (how: Rewrite, set: string, id: string) => `bench/windows/out/pivot-ours/${how}/${set}/${id}.xlsx`

/**
 * The cases of a set of files: each case of the set `from` opened from the engine's file (`how`), or from Excel's own
 * book without it (the control of what Excel's refresh makes of it), read, refreshed (`RefreshTable` of `PT1`) and read again.
 */
export function oursCases(from: string, how: Rewrite | undefined, cases: PivotCase[]): PivotCase[] {
  const edits = how === 'kept-edit' ? cases.filter((c) => c.ops.some((o) => 'edit' in o)) : cases
  return edits.map((c): PivotCase => ({ id: c.id, file: how ? oursPath(how, from, c.id) : `test/excel/win/pivot/${from}/${c.id}.xlsx`, ops: [{ read: true }, { refresh: true }] }))
}

/** A book written again by the engine. */
export async function rewritten(data: Uint8Array, how: Exclude<Rewrite, 'kept-edit'>): Promise<Uint8Array> {
  if (how === 'xlsx') return (await writeXlsx({ book: (await readXlsx(data)).book }, { values: 'omit' })).data
  return (await writeKept(data, { book: (await openKept(data)).book })).data
}

/** The book of a case's first phase with its cell edited in a session, as the case's `edit` puts its value, written over the file. */
export async function keptEdit(data: Uint8Array, cell: string, to: EditValue): Promise<Uint8Array> {
  registerPivot(pivotEngine)
  const s = createSession((await openKept(data)).book)
  const [sheet, at] = cell.split('!')
  if (to !== null && typeof to === 'object' && 'typed' in to) s.setCellText(sheet, at, to.typed)
  else {
    // Excel's refusal is the session's: the file stays as it was
    try {
      s.setCells([{ sheet, cell: at, input: to === null ? null : typeof to === 'object' ? { formula: to.formula } : { value: to } }])
    } catch {}
  }
  return (await writeKept(data, s.getSnapshot())).data
}

async function main(set: string, how: Rewrite) {
  const dir = `test/excel/win/pivot/${set}`
  const out = `bench/windows/out/pivot-ours/${how}/${set}`
  mkdirSync(out, { recursive: true })
  if (how === 'kept-edit') {
    for (const c of (await setModule<PivotSet>('pivot', set)).buildPivotCases()) {
      const e = c.ops.find((o): o is Extract<PivotCase['ops'][number], { edit: string }> => 'edit' in o)
      if (e) writeFileSync(join(out, `${c.id}.xlsx`), await keptEdit(new Uint8Array(readFileSync(join(dir, `${c.id}-1.xlsx`))), e.edit, e.to))
    }
    return
  }
  // the books a case saved at its end, not the copies of its phases (`<id>-<k>.xlsx`)
  const ids = JSON.parse(readFileSync(`${dir}.json`, 'utf8')).cases.map((c: { id: string }) => c.id) as string[]
  const books = new Set(readdirSync(dir))
  for (const id of ids) if (books.has(`${id}.xlsx`)) writeFileSync(join(out, `${id}.xlsx`), await rewritten(new Uint8Array(readFileSync(join(dir, `${id}.xlsx`))), how))
}

if (process.argv[1]?.endsWith('ours.ts')) await main(process.argv[2], (process.argv[3] ?? 'xlsx') as Rewrite)
