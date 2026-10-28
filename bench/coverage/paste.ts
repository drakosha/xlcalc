/**
 * Paste Special against Excel for Mac: `npx tsx bench/coverage/run.ts paste <set>` writes each book of `paste/<set>.ts`
 * with `writeXlsx`, Excel opens it, copies the case's range, pastes it at the target with the case's Paste Special
 * (`excel-paste.applescript`) and saves the book. The answers go to `test/excel/paste/<set>.json` and the saved files to
 * `test/excel/paste/<set>/<id>.xlsx`; `test/paste-excel.test.ts` compares them with `editStructure` without Excel.
 * `FPASTE_ONLY=id,id` runs only those cases, `FCOV_FRESH=1` every case again. `paste-predict <set>` writes what the
 * engine predicts for each case (`paste/<set>.predicted.json`) before Excel runs a held-out set.
 */
import { join } from 'node:path'
import { writeXlsx } from '../../src/xlsx.js'
import { ensureExcel, excelVersion, runSavingScript } from './excel.js'
import { pasteArgs, type PasteAnswer, type PasteCase, type PasteCases, type PasteSet } from './paste/cases.js'
import { fileAnswers, runFileCases, secondsSince, setModule, writePrediction } from './stored.js'
import { engineHas, engineOutcome } from './paste/shown.js'
import { caseFiles } from './xlsx/sent.js'
import { HERE, ROOT } from './paths.js'

const SCRIPT = join(HERE, 'excel-paste.applescript')

/** Where a set's answers go. */
export const pasteFile = (set: string) => join(ROOT, 'test/excel/paste', `${set}.json`)

/** The set's cases, the file of each case and what identifies it. */
export async function pasteSet(set: string): Promise<{ built: PasteCases; files: Map<string, Uint8Array>; sent: Record<string, string> }> {
  const built = (await setModule<PasteSet>('paste', set)).buildPasteCases()
  return { built, ...(await caseFiles(built.cases, async (c) => (await writeXlsx({ book: built.books[c.book] })).data)) }
}

/** One case in Excel: its answer and the saved file, or the repair window's text. */
async function runCase(c: PasteCase, data: Uint8Array) {
  const r = await runSavingScript(c.id, SCRIPT, data, [pasteArgs(c)], 120_000)
  if ('rejected' in r) return r
  const raw = JSON.parse(r.out) as { error: string }
  const answer: PasteAnswer = raw.error ? { error: raw.error } : {}
  return { answer, saved: r.saved }
}

/** Runs the set's cases that have no stored answer for the same case and book (FCOV_FRESH=1 runs all) and stores the answers. */
export async function stepPaste(set: string): Promise<void> {
  const { built, files, sent } = await pasteSet(set)
  const store = fileAnswers<PasteCase, PasteAnswer>(pasteFile(set), built.cases, sent)
  const t0 = Date.now()
  const save = () => store.save({ version: excelVersion(), locale: 'en-US', seconds: secondsSince(t0) }, { books: built.books, cases: built.cases })
  await ensureExcel('en-US')
  await runFileCases(`paste ${set}`, store, store.todo, 'FPASTE_ONLY', (c) => runCase(c, files.get(c.id)!), save, t0)
}

/** Writes the engine's outcome of each case of the set whose kind it has, as `test/paste-excel.test.ts` compares it. */
export async function stepPastePredict(set: string): Promise<void> {
  const { cases, books } = (await setModule<PasteSet>('paste', set)).buildPasteCases()
  const out: Record<string, unknown> = {}
  for (const c of cases) if (engineHas(c)) out[c.id] = await engineOutcome(books[c.book], c)
  writePrediction(new URL('./paste/x.ts', import.meta.url).href, set, out)
  console.log(`paste ${set}: ${Object.keys(out).length} predicted`)
}
