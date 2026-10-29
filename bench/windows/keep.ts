/**
 * Conditional formats the model does not read, in Excel for Windows (`bench/coverage/keep/rules.ts`): `keep rules-made`
 * has Excel open each book of `RULES_BOOKS` and save it (`test/excel/win/keep/rules-made.json`, the files in
 * `rules-made/`); `keep rules` does each case of the set `rules` (`RULES_SETS`) over the saved `mixed` through COM (a rule added, deleted,
 * given another priority or look; an operation of the structure) and saves it (`keep/rules.json`, `keep/rules/`). Each
 * answer is the rules the sheets list (`CondRules`: type, applies to, formula 1, priority, stop if true) and Excel's
 * error where it refused a step; `keep rules-held` and `keep rules-held-2` do the same for the held-out sets. Test: `test/xlsx-keep-rules-excel.test.ts`.
 * The files `writeKept` wrote for the Mac's sets (`KEEP_OPEN`: `open`, the edits over the made files; `struct-open`,
 * the operations of the structure; `formats-open`, the operations over the made file of formats): `keep <set>` has
 * Excel open each (COM refuses a file it would repair), compute and save it (`keep/<set>.json`, `keep/<set>/`). Test:
 * `test/xlsx-keep-win.test.ts`, against the predictions `bench/coverage/run.ts keep win-predict` wrote from the Mac's
 * answers (`bench/coverage/keep/win.predicted.json`).
 */
import { readFileSync } from 'node:fs'
import { formatsFiles, structFiles, written } from '../coverage/keep.js'
import { RULES_BOOKS, RULES_SETS, rulesFile, rulesSteps, type RulesCase } from '../coverage/keep/rules.js'
import { fileAnswers, secondsSince } from '../coverage/stored.js'
import { caseFiles } from '../coverage/xlsx/sent.js'
import { fileJobs, opSteps, stepErrors } from './books.js'
import { winFile } from './functions.js'
import { metaOf, type Step, type Tag, type WinEnv } from './remote.js'

/** What Excel answers to a case: its error where it refused a step, and each sheet's rules as the object model lists them. */
export interface RulesAnswer {
  error?: string
  rules: [sheet: string, rules: [type: number, appliesTo: string, formula: string, priority: number, stop: boolean][]][]
}

/** The Mac's sets of files `writeKept` wrote, by name: each case's file. */
export const KEEP_OPEN: Record<string, () => Promise<{ id: string; data: Uint8Array }[]>> = {
  open: written,
  'struct-open': async () => (await structFiles()).flatMap((f) => ('data' in f.written ? [{ id: f.id, data: f.written.data }] : [])),
  'formats-open': formatsFiles,
}

/** What Excel answers to a file it opened: the book's name, Excel's caption and the repair logs it wrote meanwhile. */
export interface OpenAnswer {
  opened: [name: string, caption: string, logs: string[]]
}

/** A set of `KEEP_OPEN` in Excel: each file opened, computed and saved. */
async function stepKeepOpen(what: string, tag: Tag, e: WinEnv): Promise<void> {
  const list = await KEEP_OPEN[what]()
  const data = new Map(list.map((f) => [f.id, f.data]))
  const cases = list.map((f) => ({ id: f.id }))
  const { files, sent } = await caseFiles(cases, async (c) => data.get(c.id)!)
  const store = fileAnswers<{ id: string }, OpenAnswer>(winFile(`keep/${what}.json`), cases, sent)
  const t0 = Date.now()
  const save = () => store.save({ ...metaOf(e, tag), day: new Date().toLocaleDateString('sv-SE'), seconds: secondsSince(t0) }, {})
  await fileJobs(`keep ${what}`, `keep-${what}`, store, files, () => [{ do: 'opened' }, { do: 'full' }, { do: 'saveAs' }], (_, a) => ({ opened: a.answers[0] as OpenAnswer['opened'] }), save)
}

export async function stepKeepRules(what: string, tag: Tag, e: WinEnv): Promise<void> {
  if (KEEP_OPEN[what]) return stepKeepOpen(what, tag, e)
  const made = what === 'rules-made'
  if (!made && !RULES_SETS[what]) throw new Error(`keep: rules-made | ${[...Object.keys(RULES_SETS), ...Object.keys(KEEP_OPEN)].join(' | ')}`)
  const cases: RulesCase[] = made ? Object.keys(RULES_BOOKS).map((id) => ({ id })) : RULES_SETS[what]
  const mixed = made ? undefined : new Uint8Array(readFileSync(winFile('keep/rules-made/mixed.xlsx')))
  const { files, sent } = await caseFiles(cases, async (c) => mixed ?? rulesFile(c.id))
  const store = fileAnswers<RulesCase, RulesAnswer>(winFile(`keep/${what}.json`), cases, sent)
  const t0 = Date.now()
  const save = () => store.save({ ...metaOf(e, tag), day: new Date().toLocaleDateString('sv-SE'), seconds: secondsSince(t0) }, { cases })
  const steps = (c: RulesCase): Step[] => [...rulesSteps(c, (op) => opSteps({ id: c.id, book: 'mixed', op })), { do: 'cond', kind: 'rules' }, { do: 'saveAs' }] as Step[]
  await fileJobs(`keep ${what}`, `keep-${what}`, store, files, steps, (c, a) => {
    const n = steps(c).length - 2
    const errors = stepErrors(a.answers.slice(0, n), (i) => `step ${i + 1}`)
    return { ...(errors.length ? { error: errors.join('; ') } : {}), rules: a.answers[n] as RulesAnswer['rules'] }
  }, save)
}
