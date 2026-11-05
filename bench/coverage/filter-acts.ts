/**
 * What an AutoFilter does to the rows of its sheet, against Excel for Mac: each case of `filter-acts/<set>.ts` is written
 * into a book of its own, Excel takes its steps (`excel-filter-acts.applescript`, the keyboard's handlers in
 * `excel-keys.applescript`), reading the sheet as it opened and after each, and saves the book, its file read back for the
 * filter it holds. `npx tsx bench/coverage/run.ts filter-acts <set>` stores the answers in
 * `test/excel/filter-acts/<set>.json`; `test/autofilter-excel.test.ts` compares them with the set's prediction and the
 * engine's `getAutoFilterRange`. `FACTS_ONLY=id,id` runs only those cases, `FCOV_FRESH=1` every case again.
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { declineRepair, ensureExcel, excelVersion, KEYS, runWatched, settleRejected } from './excel.js'
import { localeMeta } from './locales.js'
import { answeredLine, answerSaver, keptSame, loadStored, runCases, setModule } from './stored.js'
import { actFile, actSpec, savedFilter, type ActAnswer, type ActCase, type ActSet, type ActState } from './filter-acts/cases.js'
import { HERE, ROOT } from './paths.js'

const SCRIPT = join(HERE, 'excel-filter-acts.applescript')

export interface StoredActs {
  meta: Record<string, string | number>
  cases: ActCase[]
  answers: Record<string, ActAnswer>
}

/** The stored answers of a set: `test/excel/[win/]filter-acts/<set>.json`. */
export const actsFile = (set: string, win = false) => join(ROOT, 'test/excel', win ? 'win' : '', 'filter-acts', `${set}.json`)

/** A state as the script gives it, a list. */
export function stateOf(raw: unknown[]): ActState {
  const [range, filterMode, hidden, heights, standard, total, column, listed] = raw as [string, boolean, number[], [number, number][], number, [number, string, string] | [], string[], string[]]
  return {
    range,
    filterMode,
    hidden,
    heights: Object.fromEntries(heights.map(([r, h]) => [String(r), h])),
    standard,
    total: total.length ? { row: (total as [number, string, string])[0], text: (total as [number, string, string])[1], formula: (total as [number, string, string])[2] } : null,
    column,
    ...(listed.length ? { listed } : {}),
  }
}

/** One case in Excel. A file Excel would repair is the case's answer, as `keptOff` asks it; a window Excel shows for a step is answered by the script. */
async function runCase(c: ActCase, label: string): Promise<ActAnswer> {
  if (declineRepair()) console.log('closed a leftover book repair window')
  const data = await actFile(c)
  const written = c.keptOff ? { written: await savedFilter(data) } : {}
  return runWatched<ActAnswer>({
    script: SCRIPT,
    label: `filter-acts ${label}`,
    folder: 'ffacts',
    book: async (file) => writeFileSync(file, data),
    args: (dir) => [join(dir, 'saved.xlsx'), 'saved.xlsx', JSON.stringify(actSpec(c)), KEYS],
    timeout: 170_000,
    poll: 1000,
    after: 'case',
    fatal: false,
    repair: async (stop, open) => {
      await stop()
      await settleRejected()
      return { refused: open, ...written }
    },
    done: async (out, dir) => {
      const raw = JSON.parse(out) as { opened: unknown[]; after: unknown[][]; dialog: string }
      return { opened: stateOf(raw.opened), after: raw.after.map(stateOf), saved: await savedFilter(new Uint8Array(readFileSync(join(dir, 'saved.xlsx')))), ...written, ...(raw.dialog ? { dialog: raw.dialog } : {}) }
    },
  })
}

/** Runs the set's cases without a stored answer for the same case (FCOV_FRESH=1 runs all) and stores the answers. */
export async function stepFilterActs(set: string): Promise<void> {
  const cases = (await setModule<ActSet>('filter-acts', set)).buildActCases()
  const target = actsFile(set)
  const prev = loadStored<StoredActs>(target)
  const answers = keptSame({ cases }, prev)
  const todo = cases.filter((c) => !answers[c.id])
  const t0 = Date.now()
  const save = answerSaver(target, { version: excelVersion(), locale: localeMeta('en-US') }, true, t0, cases, answers, { was: Number(prev?.meta.seconds ?? 0) })
  await ensureExcel('en-US')
  await runCases(todo, 'FACTS_ONLY', async (c) => {
    answers[c.id] = await runCase(c, `${set} ${c.id}`)
    const a = answers[c.id]
    console.log(`${c.id}: ${a.refused ?? `opened ${a.opened?.range}; ${(a.after ?? []).map((s) => `${s.range} hidden ${s.hidden.join(',')} total ${JSON.stringify(s.total)}`).join(' / ')}; saved ${a.saved?.ref}`}${a.dialog ? `; ${a.dialog}` : ''}`)
    save()
  }, (n) => {
    save()
    console.log(answeredLine(`filter-acts ${set}`, n, answers, cases))
  })
}
