/**
 * Workbook Statistics against Excel for Mac en-US: `npx tsx bench/coverage/run.ts stats <set>` sends the file of each
 * case of `stats/<set>.ts`, Excel opens it with the case's sheet active and the script reads the rows of Review >
 * Workbook Statistics through Accessibility (`excel-stats.applescript`). The answers go to `test/excel/stats/<set>.json`;
 * `test/stats-excel.test.ts` compares them with `Session.getWorkbookStatistics`. `FSTATS_ONLY=id,id` runs only those
 * cases; `npx tsx bench/coverage/stats/predict.ts <set>` writes the engine's answers before Excel runs a set.
 */
import { writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { ensureExcel, excelVersion, runWatched } from './excel.js'
import { statsFile, statsFileOf, type StatsAnswer, type StatsCase, type StatsRow, type StatsSet } from './stats/cases.js'
import { answeredLine, answerSaver, chosen, keptSame, loadStored, setModule } from './stored.js'
import { HERE } from './paths.js'

const SCRIPT = join(HERE, 'excel-stats.applescript')

/** The window's texts as its two sections, each its rows of a label and a value, ended by its title (`Current Sheet:`, `Workbook:`). */
export function statsAnswer(out: string): StatsAnswer {
  const sections: StatsRow[][] = []
  let rows: StatsRow[] = []
  const texts = out.split('\n').map((t) => t.trim()).filter(Boolean)
  for (let i = 0; i < texts.length; i++) {
    if (texts[i].endsWith(':')) (sections.push(rows), (rows = []))
    else if (i + 1 < texts.length && !texts[i + 1].endsWith(':')) rows.push([texts[i], texts[++i]])
  }
  if (sections.length !== 2) throw new Error(`the window's sections: ${JSON.stringify(texts)}`)
  return { sheet: sections[0], workbook: sections[1] }
}

/** Runs the set's cases that have no stored answer for the same case and stores the answers. */
export async function stepStats(set: string): Promise<void> {
  const cases = (await setModule<StatsSet>('stats', set)).buildStatsCases()
  const file = statsFile(set)
  const answers = keptSame<StatsCase, StatsAnswer>({ cases }, loadStored(file))
  const t0 = Date.now()
  const save = answerSaver(file, { version: excelVersion(), locale: 'en-US' }, true, t0, cases, answers)
  const todo = cases.filter((c) => !answers[c.id]).filter(chosen('FSTATS_ONLY'))
  await ensureExcel('en-US')
  for (const c of todo) {
    const data = await statsFileOf(c)
    const out = await runWatched<string>({
      script: SCRIPT,
      label: c.id,
      folder: 'fstats',
      book: (f) => writeFileSync(f, data),
      args: () => [c.sheet],
      timeout: 120_000,
      poll: 1000,
      after: 'case',
      // the script answers the window it opens
      fatal: false,
      repair: async () => {
        throw new Error(`${c.id}: Excel would repair the book`)
      },
      done: (o) => o,
    })
    answers[c.id] = statsAnswer(out)
    save()
  }
  save()
  console.log(answeredLine(`stats ${set}`, todo.length, answers, cases))
}
