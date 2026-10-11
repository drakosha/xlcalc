/**
 * Data validation and outlines against Excel: `npx tsx bench/coverage/run.ts valid <set> [ru]` writes each book of
 * `valid/<set>.ts` with `writeXlsx`, Excel opens it, makes what the case makes, applies its operation, reads every
 * sheet's rules, values and outline over the window and saves the book (`excel-valid.applescript`). The answers go to
 * `test/excel/valid/[ru-RU/]<set>.json` and the saved files to `test/excel/valid/[ru-RU/]<set>/<id>.xlsx`;
 * `test/validation-excel.test.ts` compares them with the engine and the reader without Excel. `FVALID_ONLY=id,id` runs
 * only those cases, `FCOV_FRESH=1` every case again.
 */
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { writeXlsx } from '../../src/xlsx.js'
import { areasOf } from '../../src/validation/rules.js'
import { formatAddress } from '../../src/model.js'
import { ensureExcel, excelVersion, runSavingScript } from './excel.js'
import { localeDir } from './locales.js'
import { fileAnswers, runFileCases, secondsSince, setModule } from './stored.js'
import { opJson } from './struct.js'
import { lineRange, WINDOW, type ValidAnswer, type ValidCase, type ValidCases, type ValidMake, type ValidSet, type ValidSheetAnswer } from './valid/cases.js'
import { caseFiles } from './xlsx/sent.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
const SCRIPT = join(HERE, 'excel-valid.applescript')

export interface StoredValid extends ValidCases {
  meta: Record<string, string | number>
  /** What was sent for each case: `partsHash` of the file `writeXlsx` wrote. */
  sent: Record<string, string>
  answers: Record<string, ValidAnswer>
  /** Cases whose book Excel did not load: the repair window's text. */
  rejected?: Record<string, string>
}

/** A rule's formula as Excel's object model takes it: a literal list's items without quotes, else the formula after `=`. */
export function formulaForExcel(text: string | undefined): string {
  if (text === undefined) return ''
  const lit = /^"((?:[^"]|"")*)"$/.exec(text.trim())
  return lit ? lit[1].replace(/""/g, '"') : `=${text}`
}

const flag = (v: boolean | undefined) => (v === undefined ? '' : v ? '1' : '0')

/** A range of rules as Excel's object model takes it: areas separated by commas, whole columns and rows as cells. */
const excelRange = (range: string) => areasOf(range)!.map((x) => `${formatAddress(x.r1, x.c1)}:${formatAddress(x.r2, x.c2)}`).join(',')

/** What a case makes, as the script takes it. */
export function makeArgs(make: ValidMake[] = []): string[][] {
  return make.map((m) => {
    if ('rule' in m) {
      const r = m.rule
      const first = areasOf(m.range)![0]
      const [f1, f2] = m.excel ?? [formulaForExcel(r.formula1), formulaForExcel(r.formula2)]
      return ['rule', m.sheet, excelRange(m.range), formatAddress(first.r1, first.c1), r.type, r.errorStyle ?? '', r.operator ?? '', f1, f2 ?? '', flag(r.allowBlank), flag(r.dropdown), flag(r.showPrompt), flag(r.showError), r.promptTitle ?? '', r.prompt ?? '', r.errorTitle ?? '', r.error ?? '']
    }
    if ('summary' in m) return ['summary', m.sheet, flag(m.summary.summaryBelow), flag(m.summary.summaryRight)]
    const text = 'group' in m ? m.group : m.collapse
    return ['group' in m ? 'group' : 'collapse', m.sheet, /^\d/.test(text) ? 'rows' : 'columns', lineRange(text)]
  })
}

type RawSheet = [string, ValidSheetAnswer['rules'], ValidSheetAnswer['rows'], ValidSheetAnswer['columns'], [string, string]]


/** One case in Excel: its answer and the saved file, or the repair window's text. A window that is not a book stops the run. */
async function runCase(c: ValidCase, data: Uint8Array, window: string): Promise<{ answer: ValidAnswer; saved: Uint8Array } | { rejected: string }> {
  const r = await runSavingScript(c.id, SCRIPT, data, [JSON.stringify(makeArgs(c.make)), opJson(c), window], 150_000)
  if ('rejected' in r) return r
  const raw = JSON.parse(r.out) as { error: string; sheets: RawSheet[] }
  const answer: ValidAnswer = {
    ...(raw.error ? { error: raw.error } : {}),
    sheets: raw.sheets.map(([name, rules, rows, columns, summary]) => ({ name, rules, rows, columns, summary })),
  }
  return { answer, saved: r.saved }
}

/** Where a set's answers go: `test/excel/valid/<set>.json`, another locale's in a directory of its tag. */
export const validFile = (set: string, tag: string) => join(ROOT, 'test/excel/valid', localeDir(tag), `${set}.json`)

/** The set's cases, the file of each case and what identifies it. */
export async function validSet(set: string): Promise<{ built: ValidCases; files: Map<string, Uint8Array>; sent: Record<string, string> }> {
  const built = (await setModule<ValidSet>('valid', set)).buildValidCases()
  return { built, ...(await caseFiles(built.cases, async (c) => (await writeXlsx(built.books[c.book])).data)) }
}

/** Runs the set's cases that have no stored answer for the same case and book (FCOV_FRESH=1 runs all) and stores the answers. */
export async function stepValid(set: string, tag: string): Promise<void> {
  const { built, files, sent } = await validSet(set)
  const { books, cases } = built
  const window = built.window ?? WINDOW
  const store = fileAnswers<ValidCase, ValidAnswer>(validFile(set, tag), cases, sent)
  const t0 = Date.now()
  // the locale by its tag alone, as the step's answers have named it
  const save = () => store.save({ version: excelVersion(), locale: tag, seconds: secondsSince(t0) }, { books, cases, ...(built.window ? { window: built.window } : {}) })
  await ensureExcel(tag)
  await runFileCases(`valid ${set} (${tag})`, store, store.todo, 'FVALID_ONLY', (c) => runCase(c, files.get(c.id)!, window), save, t0)
}
