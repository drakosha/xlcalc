/**
 * The AutoFilter under what is entered on its sheet, against Excel for Mac: each case of `filter/<set>.ts` is written by
 * `writeXlsx` into a book of its own, Excel applies its criteria and takes its steps with Excel in front
 * (`excel-filter.applescript`, the keyboard's handlers in `excel-keys.applescript`), reading the AutoFilter's range
 * after each, and saves the book, its file read back by `readXlsx` for the AutoFilter it holds. `npx tsx
 * bench/coverage/run.ts filter <set>` stores the answers in `test/excel/filter/<set>.json`; `test/autofilter-excel.test.ts`
 * compares the engine with them. The clipboard's text is put back after the set. `FFILTER_ONLY=id,id` runs only those
 * cases, `FCOV_FRESH=1` every case again.
 */
import { execFileSync } from 'node:child_process'
import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { readXlsx, writeXlsx } from '../../src/xlsx.js'
import { declineRepair, ensureExcel, excelVersion, KEYS, runWatched, settleRejected } from './excel.js'
import { localeMeta } from './locales.js'
import { answeredLine, answerSaver, keptSame, loadStored, runCases, setModule } from './stored.js'
import { filterBook, type FilterAnswer, type FilterCase, type FilterSet } from './filter/cases.js'
import { HERE, ROOT } from './paths.js'

const SCRIPT = join(HERE, 'excel-filter.applescript')
/** pbpaste and pbcopy read and write text in the locale's encoding. */
const UTF8 = { ...process.env, LANG: 'en_US.UTF-8' }

export interface StoredFilter {
  meta: Record<string, string | number>
  cases: FilterCase[]
  answers: Record<string, FilterAnswer>
}

/** The stored answers of a set: `test/excel/[win/]filter/<set>.json`. */
export const filterFile = (set: string, win = false) => join(ROOT, 'test/excel', win ? 'win' : '', 'filter', `${set}.json`)

/** The AutoFilter's range a saved file holds, null for none. */
export async function savedFilter(saved: Uint8Array): Promise<string | null> {
  return (await readXlsx(saved)).book.sheets.find((s) => s.name === 'S')?.autoFilter?.range ?? null
}

/** One case in Excel. A book Excel repairs stops the run; a window Excel shows for a step is answered by the script. */
async function runCase(c: FilterCase, label: string): Promise<FilterAnswer> {
  if (declineRepair()) console.log('closed a leftover book repair window')
  return runWatched<FilterAnswer>({
    script: SCRIPT,
    label: `filter ${label}`,
    folder: 'ffilter',
    book: async (file) => {
      const { data, issues } = await writeXlsx({ book: filterBook(c) }, { values: 'omit' })
      if (issues.length) throw new Error(`filter ${label}: the book does not go into a file whole: ${JSON.stringify(issues)}`)
      writeFileSync(file, data)
    },
    args: (dir) => [join(dir, 'saved.xlsx'), 'saved.xlsx', JSON.stringify({ filter: c.filter, ...(c.criteria ? { criteria: c.criteria } : {}), steps: c.steps }), KEYS],
    timeout: 170_000,
    poll: 1000,
    after: 'case',
    fatal: false,
    repair: async (stop) => {
      await stop()
      await settleRejected()
      throw new Error(`filter ${label}: Excel would repair the book`)
    },
    done: async (out, dir) => {
      const raw = JSON.parse(out) as { after: string[]; dialog: string }
      return { after: raw.after, saved: await savedFilter(new Uint8Array(readFileSync(join(dir, 'saved.xlsx')))), ...(raw.dialog ? { dialog: raw.dialog } : {}) }
    },
  })
}

/** Runs the set's cases without a stored answer for the same case (FCOV_FRESH=1 runs all) and stores the answers; the clipboard's text is put back after. */
export async function stepFilter(set: string): Promise<void> {
  const cases = (await setModule<FilterSet>('filter', set)).buildFilterCases()
  const target = filterFile(set)
  const prev = loadStored<StoredFilter>(target)
  const answers = keptSame({ cases }, prev)
  const todo = cases.filter((c) => !answers[c.id])
  const t0 = Date.now()
  const save = answerSaver(target, { version: excelVersion(), locale: localeMeta('en-US') }, true, t0, cases, answers, { was: Number(prev?.meta.seconds ?? 0) })
  await ensureExcel('en-US')
  const clipboard = execFileSync('pbpaste', { encoding: 'utf8', env: UTF8 })
  try {
    await runCases(todo, 'FFILTER_ONLY', async (c) => {
      answers[c.id] = await runCase(c, `${set} ${c.id}`)
      console.log(`${c.id}: ${answers[c.id].after.join(', ')}; saved ${answers[c.id].saved}${answers[c.id].dialog ? `; ${answers[c.id].dialog}` : ''}`)
      save()
    }, (n) => {
      save()
      console.log(answeredLine(`filter ${set}`, n, answers, cases))
    })
  } finally {
    execFileSync('pbcopy', { input: clipboard, env: UTF8 })
  }
}
