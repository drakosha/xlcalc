/**
 * Hidden sheets against Excel en-US: `npx tsx bench/coverage/run.ts hidden <set>` writes each book of `hidden/<set>.ts`
 * with `writeXlsx` (changed by hand for a raw case), Excel opens it, sets the visibility of sheets, applies the case's
 * sheet operation, reads each sheet's visibility, the active sheet and the read cells' values, and saves the book
 * (`excel-hidden.applescript`). The answers go to `test/excel/hidden/<set>.json` and the saved files to
 * `test/excel/hidden/<set>/<id>.xlsx`; `test/hidden-sheets-excel.test.ts` compares them with the engine and the reader
 * without Excel. A book Excel would repair is recorded in `rejected`. `FHIDDEN_ONLY=id,id` runs only those cases,
 * `FCOV_FRESH=1` every case again.
 */
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { ensureExcel, excelVersion, runSavingScript } from './excel.js'
import { hiddenFile, type HiddenAnswer, type HiddenCase, type HiddenCases, type HiddenSet } from './hidden/cases.js'
import { localeMeta } from './locales.js'
import { fileAnswers, runFileCases, secondsSince, setModule } from './stored.js'
import { opJson } from './struct.js'
import { caseFiles } from './xlsx/sent.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
const SCRIPT = join(HERE, 'excel-hidden.applescript')

export interface StoredHidden extends HiddenCases {
  meta: Record<string, string | number>
  /** What was sent for each case: `partsHash` of its file. */
  sent: Record<string, string>
  answers: Record<string, HiddenAnswer>
  /** Cases whose book Excel did not load: why. */
  rejected?: Record<string, string>
}

/** The set's cases, the file of each case and what identifies it. */
export async function hiddenSet(set: string): Promise<{ built: HiddenCases; files: Map<string, Uint8Array>; sent: Record<string, string> }> {
  const built = (await setModule<HiddenSet>('hidden', set)).buildHiddenCases()
  return { built, ...(await caseFiles(built.cases, (c) => hiddenFile(c, built))) }
}

/** One case in Excel: its answer and the saved file, or the repair window's text. A window that is not a book stops the run. */
async function runCase(c: HiddenCase, data: Uint8Array, read: HiddenCases['read']): Promise<{ answer: HiddenAnswer; saved: Uint8Array } | { rejected: string }> {
  const makes = (c.make ?? []).map((m) => [m.sheet, m.visible])
  const r = await runSavingScript(c.id, SCRIPT, data, [JSON.stringify(makes), opJson(c), read.sheet, JSON.stringify(read.cells)], 120_000)
  if ('rejected' in r) return r
  const { error, ...rest } = JSON.parse(r.out) as Omit<HiddenAnswer, 'error'> & { error: string }
  return { answer: { ...(error ? { error } : {}), ...rest, values: rest.values.map((v) => (v === '' ? null : v)) }, saved: r.saved }
}

/** Runs the set's cases that have no stored answer for the same case and file (FCOV_FRESH=1 runs all) and stores the answers. */
export async function stepHidden(set: string): Promise<void> {
  const { built, files, sent } = await hiddenSet(set)
  const store = fileAnswers<HiddenCase, HiddenAnswer>(join(ROOT, `test/excel/hidden/${set}.json`), built.cases, sent)
  const t0 = Date.now()
  const save = () => store.save({ version: excelVersion(), locale: localeMeta('en-US'), seconds: secondsSince(t0) }, built)
  await ensureExcel('en-US')
  await runFileCases(`hidden ${set}`, store, store.todo, 'FHIDDEN_ONLY', (c) => runCase(c, files.get(c.id)!, built.read), save, t0)
}
