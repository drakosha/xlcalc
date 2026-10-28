/**
 * Paste Special against Excel for Mac: `npx tsx bench/coverage/run.ts paste <set>` writes each book of `paste/<set>.ts`
 * with `writeXlsx`, Excel opens it, copies the case's range, pastes it at the target with the case's Paste Special
 * (`excel-paste.applescript`) and saves the book. The answers go to `test/excel/paste/<set>.json` and the saved files to
 * `test/excel/paste/<set>/<id>.xlsx` (`paste <set> ru`: Excel ru-RU, into `test/excel/paste/ru-RU/`), with the formula
 * in the user's language of each cell a case names (`local`); `test/paste-excel.test.ts` compares them with
 * `editStructure` without Excel.
 * `FPASTE_ONLY=id,id` runs only those cases, `FCOV_FRESH=1` every case again. `paste-predict <set>` writes what the
 * engine predicts for each case (`paste/<set>.predicted.json`) before Excel runs a held-out set.
 */
import { join } from 'node:path'
import { writeXlsx } from '../../src/xlsx.js'
import { ensureExcel, excelVersion, runSavingScript } from './excel.js'
import { pasteArgs, WINDOWS_ONLY, type PasteAnswer, type PasteCase, type PasteCases, type PasteSet } from './paste/cases.js'
import { fileAnswers, runFileCases, secondsSince, setModule, writePrediction } from './stored.js'
import { engineHas, engineOutcome } from './paste/shown.js'
import { caseFiles, editParts } from './xlsx/sent.js'
import { HERE, ROOT } from './paths.js'
import { localeDir } from './locales.js'

const SCRIPT = join(HERE, 'excel-paste.applescript')

/** Where a set's answers go, a locale's other than en-US in a directory of its tag. */
export const pasteFile = (set: string, tag = 'en-US') => join(ROOT, 'test/excel/paste', `${localeDir(tag)}${set}.json`)

/** The formulas of `cells` on each sheet as Excel saves a formula typed into a cell: `<f>`, without the dynamic array's metadata. */
const plainFormulas = (cells: string[]) => (name: string, xml: string) =>
  name.startsWith('xl/worksheets/') ? xml.replace(/<c r="([A-Z]+\d+)"([^>]*?) cm="1"><f t="array" ref="\1">/g, (all, at: string, rest: string) => (cells.includes(at) ? `<c r="${at}"${rest}><f>` : all)) : undefined

/** The set's cases, the file of each case and what identifies it. */
export async function pasteSet(set: string): Promise<{ built: PasteCases; files: Map<string, Uint8Array>; sent: Record<string, string> }> {
  const built = (await setModule<PasteSet>('paste', set)).buildPasteCases()
  return {
    built,
    ...(await caseFiles(built.cases, async (c) => {
      const { data } = await writeXlsx({ book: built.books[c.book] })
      const plain = built.plain?.[c.book]
      return plain ? editParts(data, plainFormulas(plain)) : data
    })),
  }
}

/** One case in Excel: its answer and the saved file, or the repair window's text. */
async function runCase(c: PasteCase, data: Uint8Array) {
  const r = await runSavingScript(c.id, SCRIPT, data, [pasteArgs(c)], 120_000)
  if ('rejected' in r) return r
  const raw = JSON.parse(r.out) as { error: string; local: string[] }
  const answer = pasteAnswer(c, raw.error, raw.local)
  return { answer, saved: r.saved }
}

/** Runs the set's cases that have no stored answer for the same case and book (FCOV_FRESH=1 runs all) and stores the answers. */
/** A case's answer: Excel's error, and the formulas it read in the user's language by cell. */
export const pasteAnswer = (c: PasteCase, error: string | undefined, local: string[]): PasteAnswer => ({
  ...(error ? { error } : {}),
  ...(c.local?.length ? { local: Object.fromEntries(c.local.map((a, i) => [a, local[i]])) } : {}),
})

export async function stepPaste(set: string, tag = 'en-US'): Promise<void> {
  const { built: all, files, sent } = await pasteSet(set)
  const built = { ...all, cases: all.cases.filter((c) => !WINDOWS_ONLY.includes(c.what)) }
  if (!built.cases.length) return void console.log(`paste ${set}: no case for Excel for Mac`)
  const store = fileAnswers<PasteCase, PasteAnswer>(pasteFile(set, tag), built.cases, sent)
  const t0 = Date.now()
  const save = () => store.save({ version: excelVersion(), locale: tag, seconds: secondsSince(t0) }, { books: built.books, cases: built.cases })
  await ensureExcel(tag)
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
