/**
 * Conditional formats the model does not read, in Excel for Windows (`bench/coverage/keep/rules.ts`): `keep rules-made`
 * has Excel open each book of `RULES_BOOKS` and save it (`test/excel/win/keep/rules-made.json`, the files in
 * `rules-made/`); `keep rules` does each case of `RULES_CASES` over the saved `mixed` through COM (a rule added, deleted,
 * given another priority or look; an operation of the structure) and saves it (`keep/rules.json`, `keep/rules/`). Each
 * answer is the rules the sheets list (`CondRules`: type, applies to, formula 1, priority, stop if true) and Excel's
 * error where it refused a step; `keep rules-held` does the same for `RULES_HELD`. Test: `test/xlsx-keep-rules-excel.test.ts`.
 */
import { readFileSync } from 'node:fs'
import { RULES_BOOKS, RULES_CASES, RULES_HELD, rulesFile, rulesSteps, type RulesCase } from '../coverage/keep/rules.js'
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

export async function stepKeepRules(what: string, tag: Tag, e: WinEnv): Promise<void> {
  if (what !== 'rules-made' && what !== 'rules' && what !== 'rules-held') throw new Error('keep: rules-made | rules | rules-held')
  const made = what === 'rules-made'
  const cases: RulesCase[] = made ? Object.keys(RULES_BOOKS).map((id) => ({ id })) : what === 'rules' ? RULES_CASES : RULES_HELD
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
