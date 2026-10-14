/**
 * Notes, hyperlinks and frozen panes against Excel en-US: `npx tsx bench/coverage/run.ts parts <set>` writes each book
 * of `parts/<set>.ts` with `writeXlsx` (changed by hand for a raw case), Excel opens it, makes what the case makes, applies its operation, reads every
 * sheet's parts and saves the book (`excel-parts.applescript`). The answers go to `test/excel/parts/<set>.json` and the
 * saved files to `test/excel/parts/<set>/<id>.xlsx`; `test/sheet-parts-excel.test.ts` compares them with the writer,
 * the reader and `editStructure` without Excel. `FPARTS_ONLY=id,id` runs only those cases, `FCOV_FRESH=1` every case again.
 */
import { join } from 'node:path'
import { writeXlsx } from '../../src/xlsx.js'
import { ensureExcel, excelVersion, runSavingScript } from './excel.js'
import { localeMeta } from './locales.js'
import type { PartsAnswer, PartsCase, PartsCases, PartsMake, PartsSet } from './parts/cases.js'
import { fileAnswers, runFileCases, secondsSince, setModule } from './stored.js'
import { threadedFile } from './parts/threaded.js'
import { opJson } from './struct.js'
import { caseFiles, rawFile } from './xlsx/sent.js'
import { HERE, ROOT } from './paths.js'

const SCRIPT = join(HERE, 'excel-parts.applescript')

export interface StoredParts extends PartsCases {
  meta: Record<string, string | number>
  /** What was sent for each case: `partsHash` of the file `writeXlsx` wrote. */
  sent: Record<string, string>
  answers: Record<string, PartsAnswer>
  /** Cases whose book Excel did not load: the repair window's text. */
  rejected?: Record<string, string>
}

/** What a case makes, as the script takes it. */
export function makeArgs(make: PartsMake[] = []): (string | number)[][] {
  return make.map((m) => {
    if ('note' in m) return ['note', m.sheet, m.note, m.text]
    if ('link' in m) return ['link', m.sheet, m.link, m.url ?? '', m.location ?? '', m.tooltip ?? '']
    return ['freeze', m.freeze, m.rows, m.columns]
  })
}

/** The file a case's book is written to, the same on both platforms: by `writeXlsx`, with the book's threads put in, changed by hand for a raw case. */
export async function partsFile(c: PartsCase, set: PartsCases): Promise<Uint8Array> {
  const threads = set.threads?.[c.book]
  return rawFile(threads ? await threadedFile(set.books[c.book], threads) : (await writeXlsx({ book: set.books[c.book] })).data, c.raw)
}

/** The set's cases, the file of each case and what identifies it. */
export async function partsSet(set: string): Promise<{ built: PartsCases; files: Map<string, Uint8Array>; sent: Record<string, string> }> {
  const built = (await setModule<PartsSet>('parts', set)).buildPartsCases()
  return { built, ...(await caseFiles(built.cases, (c) => partsFile(c, built))) }
}

type RawSheet = [string, [string, string, string][], [string, string, string, string, string][], [boolean, number, number, number, number, number, number]]

const plain = (a: string) => a.replaceAll('$', '')

/** One case in Excel: its answer and the saved file, or the repair window's text. A window that is not a book stops the run. */
async function runCase(c: PartsCase, data: Uint8Array): Promise<{ answer: PartsAnswer; saved: Uint8Array } | { rejected: string }> {
  const r = await runSavingScript(c.id, SCRIPT, data, [JSON.stringify(makeArgs(c.make)), opJson(c)], 120_000)
  if ('rejected' in r) return r
  const raw = JSON.parse(r.out) as { error: string; sheets: RawSheet[] }
  const answer: PartsAnswer = {
    ...(raw.error ? { error: raw.error } : {}),
    sheets: raw.sheets.map(([name, notes, links, [freeze, splitRow, splitColumn, , , scrollRow, scrollColumn]]) => ({
      name,
      notes: Object.fromEntries(notes.map(([a, author, text]) => [plain(a), { author, text }])),
      links: links.map(([range, address, subAddress, tip, display]) => ({ range: plain(range), address, subAddress, tip, display })),
      panes: { freeze, splitRow, splitColumn, scrollRow, scrollColumn },
    })),
  }
  return { answer, saved: r.saved }
}

/** Runs the set's cases that have no stored answer for the same case and book (FCOV_FRESH=1 runs all) and stores the answers. */
export async function stepParts(set: string): Promise<void> {
  const { built, files, sent } = await partsSet(set)
  const { books, cases, threads } = built
  const store = fileAnswers<PartsCase, PartsAnswer>(join(ROOT, `test/excel/parts/${set}.json`), cases, sent)
  const t0 = Date.now()
  const save = () => store.save({ version: excelVersion(), locale: localeMeta('en-US'), seconds: secondsSince(t0) }, { books, cases, ...(threads ? { threads } : {}) })
  await ensureExcel('en-US')
  await runFileCases(`parts ${set}`, store, store.todo, 'FPARTS_ONLY', (c) => runCase(c, files.get(c.id)!), save, t0)
}
