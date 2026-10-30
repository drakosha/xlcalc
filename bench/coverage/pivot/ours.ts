/**
 * The engine's files of the books of a pivot set, for the sets that open them in Excel (`write-1`, a case's `file`):
 * each book Excel saved for a case of the set (`test/excel/win/pivot/<set>/<id>.xlsx`) read and written again, by
 * `writeXlsx` (`xlsx`) or by `openKept` and `writeKept` as it is (`kept`), into `bench/windows/out/pivot-ours/<how>/<set>/`.
 *
 *   npx tsx bench/coverage/pivot/ours.ts basic xlsx
 */
import { mkdirSync, readFileSync, readdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import type { PivotCase } from './cases.js'
import { openKept, writeKept } from '../../../src/xlsx-keep.js'
import { readXlsx, writeXlsx } from '../../../src/xlsx.js'

export type Rewrite = 'xlsx' | 'kept'

/** Where the engine's file of a case's book is, from the package's folder. */
export const oursPath = (how: Rewrite, set: string, id: string) => `bench/windows/out/pivot-ours/${how}/${set}/${id}.xlsx`

/**
 * The cases of a set of files: each case of the set `from` opened from the engine's file (`how`), or from Excel's own
 * book without it (the control of what Excel's refresh makes of it), read, refreshed (`RefreshTable` of `PT1`) and read again.
 */
export function oursCases(from: string, how: Rewrite | undefined, cases: PivotCase[]): PivotCase[] {
  return cases.map((c): PivotCase => ({ id: c.id, file: how ? oursPath(how, from, c.id) : `test/excel/win/pivot/${from}/${c.id}.xlsx`, ops: [{ read: true }, { refresh: true }] }))
}

/** A book written again by the engine. */
export async function rewritten(data: Uint8Array, how: Rewrite): Promise<Uint8Array> {
  if (how === 'xlsx') return (await writeXlsx({ book: (await readXlsx(data)).book }, { values: 'omit' })).data
  return (await writeKept(data, { book: (await openKept(data)).book })).data
}

async function main(set: string, how: Rewrite) {
  const dir = `test/excel/win/pivot/${set}`
  const out = `bench/windows/out/pivot-ours/${how}/${set}`
  mkdirSync(out, { recursive: true })
  // the books a case saved at its end, not the copies of its phases (`<id>-<k>.xlsx`)
  const ids = JSON.parse(readFileSync(`${dir}.json`, 'utf8')).cases.map((c: { id: string }) => c.id) as string[]
  const books = new Set(readdirSync(dir))
  for (const id of ids) if (books.has(`${id}.xlsx`)) writeFileSync(join(out, `${id}.xlsx`), await rewritten(new Uint8Array(readFileSync(join(dir, `${id}.xlsx`))), how))
}

if (process.argv[1]?.endsWith('ours.ts')) await main(process.argv[2], (process.argv[3] ?? 'xlsx') as Rewrite)
