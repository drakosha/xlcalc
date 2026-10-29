/**
 * CSV files against Excel: `npx tsx bench/coverage/run.ts csv <set>` writes each case of `csv/<set>.ts` (a CSV file
 * under its name, or a book as xlsx) into Excel's sandbox; Excel, in the set's locale, opens it, saves it as xlsx, then
 * as CSV (`excel-csv.applescript`). The answers (the sheet's name, the CSV as base64) go to
 * `test/excel/csv/[<tag>/]<set>.json` with the day of the run, the saved xlsx files beside them;
 * `test/csv-excel.test.ts` compares `readCsv` with what Excel read and `writeCsv` with what it saved.
 * `FCSV_ONLY=id,id` runs only those cases, `FCOV_FRESH=1` every case again.
 */
import { createHash } from 'node:crypto'
import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { caseFile, fileName, type CsvAnswer, type CsvCase, type CsvSet } from './csv/cases.js'
import { ensureExcel, excelVersion, runWatched } from './excel.js'
import { localeDir, localeMeta } from './locales.js'
import { HERE, ROOT } from './paths.js'
import { fileAnswers, runFileCases, secondsSince, setModule, type FileAnswer } from './stored.js'
import { partsHash } from './xlsx/sent.js'

const SCRIPT = join(HERE, 'excel-csv.applescript')

/** The set's cases, the file of each case and what identifies it. */
export async function csvSet(set: string): Promise<{ locale: string; cases: CsvCase[]; files: Map<string, Uint8Array>; sent: Record<string, string> }> {
  const built = (await setModule<CsvSet>('csv', set)).buildCsvCases()
  const files = new Map(await Promise.all(built.cases.map(async (c) => [c.id, await caseFile(c)] as const)))
  // a book by its parts, as the other steps send it; a CSV file by its bytes
  const hash = async (c: CsvCase) => (c.book ? partsHash(files.get(c.id)!) : createHash('sha256').update(files.get(c.id)!).update(fileName(c)).digest('hex'))
  const sent = Object.fromEntries(await Promise.all(built.cases.map(async (c) => [c.id, await hash(c)] as const)))
  return { locale: built.locale ?? 'en-US', cases: built.cases, files, sent }
}

/** The stored answers' file of a set. */
export const csvFile = (set: string, tag: string): string => join(ROOT, `test/excel/csv/${localeDir(tag)}${set}.json`)

/** One case in Excel: the sheet it named and the CSV and xlsx it saved. A window that is not a book stops the run. */
function runCase(c: CsvCase, data: Uint8Array): Promise<FileAnswer<CsvAnswer>> {
  return runWatched<FileAnswer<CsvAnswer>>({
    script: SCRIPT,
    label: c.id,
    folder: 'fcsv',
    file: fileName(c),
    book: (file) => writeFileSync(file, data),
    args: (dir) => [join(dir, 'saved.xlsx'), 'saved.xlsx', join(dir, 'saved.csv'), 'saved.csv'],
    timeout: 120_000,
    poll: 1000,
    after: 'case',
    fatal: true,
    repair: async () => {
      throw new Error(`${c.id}: Excel would repair the file`)
    },
    done: (out, dir) => ({
      answer: { sheet: out.trim(), csv: readFileSync(join(dir, 'saved.csv')).toString('base64') },
      saved: new Uint8Array(readFileSync(join(dir, 'saved.xlsx'))),
    }),
  })
}

/** Runs the set's cases that have no stored answer for the same case and file and stores the answers. */
export async function stepCsv(set: string): Promise<void> {
  const { locale, cases, files, sent } = await csvSet(set)
  const store = fileAnswers<CsvCase, CsvAnswer>(csvFile(set, locale), cases, sent)
  const t0 = Date.now()
  const save = () => store.save({ version: excelVersion(), locale: localeMeta(locale), date: new Date().toISOString().slice(0, 10), seconds: secondsSince(t0) }, { cases })
  await ensureExcel(locale)
  await runFileCases(`csv ${set}`, store, store.todo, 'FCSV_ONLY', (c) => runCase(c, files.get(c.id)!), save, t0)
}
