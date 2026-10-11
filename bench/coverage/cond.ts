/**
 * Conditional formats against Excel for Mac: `npx tsx bench/coverage/run.ts cond <set>` writes each book of
 * `cond/<set>.ts` with `writeXlsx` (changed by hand for a raw case), Excel opens it in the set's locale, applies the
 * case's operation, computes the book, reads what the listed cells show by their display format and saves the book
 * (`excel-cond.applescript`). The answers go to `test/excel/cond/<set>.json` (another locale's to `cond/<tag>/`) and the
 * saved files beside them;
 * `test/cond-excel.test.ts` compares them with the engine and the reader without Excel. Cases whose rules Excel makes
 * through its object model are Windows's alone (`bench/windows/cond.ts`): Excel for Mac's `make new format condition`
 * stopped Excel. `FCOND_ONLY=id,id` runs only those cases, `FCOV_FRESH=1` every case again.
 */
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { condFile, fromRgb, type CondAnswer, type CondCase, type CondCases, type CondSet, type Look } from './cond/cases.js'
import { ensureExcel, excelVersion, runSavingScript } from './excel.js'
import { localeDir, localeMeta } from './locales.js'
import { fileAnswers, runFileCases, secondsSince, setModule } from './stored.js'
import { opJson } from './struct.js'
import { caseFiles } from './xlsx/sent.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
const SCRIPT = join(HERE, 'excel-cond.applescript')

export interface StoredCond extends CondCases {
  meta: Record<string, string | number>
  /** What was sent for each case: `partsHash` of the file `writeXlsx` wrote. */
  sent: Record<string, string>
  answers: Record<string, CondAnswer>
  /** Cases whose book Excel did not load. */
  rejected?: Record<string, string>
}

/** The set's cases and books, and where its answers go (`test/excel/<base>cond/[<tag>/]<set>`). */
export async function condSet(set: string): Promise<{ built: CondCases; tag: string; files: Map<string, Uint8Array>; sent: Record<string, string> }> {
  const built = (await setModule<CondSet>('cond', set)).buildCondCases()
  return { built, tag: built.locale ?? 'en-US', ...(await caseFiles(built.cases, (c) => condFile(built, c))) }
}

type RawLook = [address: string, text: string, fill: number[], colorIndex: string, font: number[], bold: boolean, italic: boolean]

/** One case in Excel: its answer and the saved file, or the repair window's text. A window that is not a book stops the run. */
async function runCase(c: CondCase, data: Uint8Array): Promise<{ answer: CondAnswer; saved: Uint8Array } | { rejected: string }> {
  const r = await runSavingScript(c.id, SCRIPT, data, [opJson(c), JSON.stringify(c.read)])
  if ('rejected' in r) return r
  const raw = JSON.parse(r.out) as { error: string; looks: [string, RawLook[]][] }
  const looks: CondAnswer['looks'] = {}
  for (const [sheet, cells] of raw.looks)
    looks[sheet] = Object.fromEntries(cells.map(([a, text, fill, ci, font, bold, italic]): [string, Look] => [a.replaceAll('$', ''), { text, fill: ci === 'color index none' ? null : fromRgb(fill), font: fromRgb(font), bold, italic }]))
  return { answer: { ...(raw.error ? { error: raw.error } : {}), looks }, saved: r.saved }
}

/** Runs the set's cases (but those whose rules Excel makes) that have no stored answer for the same case and file, and stores the answers. */
export async function stepCond(set: string): Promise<void> {
  const { built, tag, files, sent } = await condSet(set)
  const { books, cases } = built
  const store = fileAnswers<CondCase, CondAnswer>(join(ROOT, `test/excel/cond/${localeDir(tag)}${set}.json`), cases, sent)
  const todo = store.todo.filter((c) => !c.make)
  if (!todo.length) return console.log(`cond ${set}: nothing to run on the Mac`)
  const t0 = Date.now()
  const day = new Date().toLocaleDateString('sv-SE')
  const save = () => store.save({ version: excelVersion(), locale: localeMeta(tag), day, seconds: secondsSince(t0) }, { books, cases, ...(built.locale ? { locale: built.locale } : {}) })
  await ensureExcel(tag)
  await runFileCases(`cond ${set}`, store, todo, 'FCOND_ONLY', (c) => runCase(c, files.get(c.id)!), save, t0)
}
