/**
 * Paste Special against Excel for Mac: `npx tsx bench/coverage/run.ts paste <set>` writes each book of `paste/<set>.ts`
 * with `writeXlsx`, Excel opens it, copies the case's range, pastes it at the target with the case's Paste Special
 * (`excel-paste.applescript`) and saves the book. The answers go to `test/excel/paste/<set>.json` and the saved files to
 * `test/excel/paste/<set>/<id>.xlsx` (`paste <set> ru`: Excel ru-RU, into `test/excel/paste/ru-RU/`), with the formula
 * in the user's language of each cell a case names (`local`); `test/paste-excel.test.ts` compares them with
 * `editStructure` without Excel. All Merging Conditional Formats, which AppleScript does not name, is pasted through
 * Excel's Paste Special window (Command+Control+V, its buttons clicked through System Events); a case that copies from
 * another workbook (`fromBook`) has that book open beside the target's.
 * `FPASTE_ONLY=id,id` runs only those cases, `FCOV_FRESH=1` every case again. `paste-predict <set>` writes what the
 * engine predicts for each case (`paste/<set>.predicted.json`) before Excel runs a held-out set.
 */
import { join } from 'node:path'
import { writeXlsx } from '../../src/xlsx.js'
import { ensureExcel, excelVersion, runSavingScript } from './excel.js'
import { pasteArgs, viaWindow, type PasteAnswer, type PasteCase, type PasteCases, type PasteSet } from './paste/cases.js'
import { fileAnswers, runFileCases, secondsSince, setModule, writePrediction } from './stored.js'
import { engineHas, engineOutcome, localFormulas } from './paste/shown.js'
import { caseFiles, editParts, partsHash } from './xlsx/sent.js'
import { HERE, ROOT } from './paths.js'
import { localeDir } from './locales.js'

const SCRIPT = join(HERE, 'excel-paste.applescript')

/** Where a set's answers go, a locale's other than en-US in a directory of its tag. */
export const pasteFile = (set: string, tag = 'en-US') => join(ROOT, 'test/excel/paste', `${localeDir(tag)}${set}.json`)

/** The formulas of `cells` on each sheet as Excel saves a formula typed into a cell: `<f>`, without the dynamic array's metadata. */
const plainFormulas = (cells: string[]) => (name: string, xml: string) =>
  name.startsWith('xl/worksheets/') ? xml.replace(/<c r="([A-Z]+\d+)"([^>]*?) cm="1"><f t="array" ref="\1">/g, (all, at: string, rest: string) => (cells.includes(at) ? `<c r="${at}"${rest}><f>` : all)) : undefined

/** A book's file as the step sends it: written, its plain formulas and its changes by hand made. */
async function bookFile(built: PasteCases, key: string): Promise<Uint8Array> {
  const { data } = await writeXlsx({ book: built.books[key] })
  const plain = built.plain?.[key]
  const file = plain ? await editParts(data, plainFormulas(plain)) : data
  return built.made?.[key] ? built.made[key](file) : file
}

/** The name of the other book a case copies from (`fromBook`), beside the target's in Excel's folder. */
export const SOURCE_FILE = 'source.xlsx'

/** The set's cases, the file of each case and of the other book it copies from, and what identifies them. */
export async function pasteSet(set: string): Promise<{ built: PasteCases; files: Map<string, Uint8Array>; sources: Map<string, Uint8Array>; sent: Record<string, string> }> {
  const built = (await setModule<PasteSet>('paste', set)).buildPasteCases()
  const sources = new Map(await Promise.all(built.cases.flatMap((c) => (c.fromBook ? [(async () => [c.id, await bookFile(built, c.fromBook!)] as const)()] : []))))
  const { files, sent } = await caseFiles(built.cases, (c) => bookFile(built, c.book))
  for (const [id, data] of sources) sent[id] += ` ${await partsHash(data)}`
  return { built, files, sources, sent }
}

/** One case in Excel: its answer and the saved file, or the repair window's text. Paste Special's window, which the script answers, stops no run. */
async function runCase(c: PasteCase, data: Uint8Array, source: Uint8Array | undefined) {
  const r = await runSavingScript(c.id, SCRIPT, data, [pasteArgs(c)], 120_000, { beside: source ? { [SOURCE_FILE]: source } : {}, fatal: !viaWindow(c) })
  if ('rejected' in r) return r
  const raw = JSON.parse(r.out) as { error: string; local: string[] }
  const answer = pasteAnswer(c, raw.error, raw.local)
  return { answer, saved: r.saved }
}

/** A case's answer: Excel's error, and the formulas it read in the user's language by cell. */
export const pasteAnswer = (c: PasteCase, error: string | undefined, local: string[]): PasteAnswer => ({
  ...(error ? { error } : {}),
  ...(c.local?.length ? { local: Object.fromEntries(c.local.map((a, i) => [a, local[i]])) } : {}),
})

/** Runs the set's cases that have no stored answer for the same case and book (FCOV_FRESH=1 runs all) and stores the answers. */
export async function stepPaste(set: string, tag = 'en-US'): Promise<void> {
  const { built, files, sources, sent } = await pasteSet(set)
  const store = fileAnswers<PasteCase, PasteAnswer>(pasteFile(set, tag), built.cases, sent)
  const t0 = Date.now()
  const save = () => store.save({ version: excelVersion(), locale: tag, seconds: secondsSince(t0) }, { books: built.books, cases: built.cases })
  await ensureExcel(tag)
  await runFileCases(`paste ${set}`, store, store.todo, 'FPASTE_ONLY', (c) => runCase(c, files.get(c.id)!, sources.get(c.id)), save, t0)
}

/** Writes the engine's outcome of each case of the set whose kind it has, as `test/paste-excel.test.ts` compares it, and the set's own prediction of a case from another workbook (`carried`). */
export async function stepPastePredict(set: string): Promise<void> {
  const mod = await setModule<PasteSet>('paste', set)
  const { cases, books } = mod.buildPasteCases()
  const out: Record<string, unknown> = {}
  for (const c of cases) {
    if (c.fromBook && mod.carried) out[c.id] = { shown: await mod.carried(c) }
    if (!engineHas(c)) continue
    const r = await engineOutcome(books[c.book], c)
    out[c.id] = c.local?.length && 'shown' in r ? { ...r, local: localFormulas(r.shown, c, books[c.book].locale ?? 'en-US') } : r
  }
  writePrediction(new URL('./paste/x.ts', import.meta.url).href, set, out)
  console.log(`paste ${set}: ${Object.keys(out).length} predicted`)
}
