/**
 * Fields of a workbook against Excel for Mac en-US: `npx tsx bench/coverage/run.ts fields <set>` sends the file of each
 * case of `fields/<set>.ts` (`fieldsFileOf`, as Excel for Windows gets it), Excel opens it, reads each sheet's window
 * and saves the book (`excel-fields.applescript`). The answers go to `test/excel/fields/<set>.json` and the saved files
 * to `test/excel/fields/<set>/<id>.xlsx`; `test/fields-excel.test.ts` compares them with the workbook. `FFIELDS_ONLY=id,id`
 * runs only those cases.
 */
import { join } from 'node:path'
import { ensureExcel, excelVersion, runSavingScript } from './excel.js'
import { fieldsFile, fieldsFileOf, panesAnswer, type FieldsCase, type FieldsSet, type PanesAnswer, type RawPanes } from './fields/cases.js'
import { fileAnswers, runFileCases, secondsSince, setModule } from './stored.js'
import { caseFiles } from './xlsx/sent.js'
import { HERE } from './paths.js'

const SCRIPT = join(HERE, 'excel-fields.applescript')

/** What Excel for Mac reads of a case: each sheet's window. */
export interface MacFieldsAnswer {
  panes: PanesAnswer[]
}

async function runCase(c: FieldsCase, data: Uint8Array) {
  const r = await runSavingScript(c.id, SCRIPT, data, [])
  if ('rejected' in r) return r
  return { answer: { panes: (JSON.parse(r.out) as RawPanes[]).map(panesAnswer) }, saved: r.saved }
}

/** Runs the set's cases that have no stored answer for the same case and file and stores the answers. */
export async function stepFields(set: string): Promise<void> {
  const cases = (await setModule<FieldsSet>('fields', set)).buildFieldsCases()
  const { files, sent } = await caseFiles(cases, fieldsFileOf)
  const store = fileAnswers<FieldsCase, MacFieldsAnswer>(fieldsFile(set, 'fields'), cases, sent)
  const t0 = Date.now()
  const save = () => store.save({ version: excelVersion(), locale: 'en-US', seconds: secondsSince(t0) }, { cases })
  await ensureExcel('en-US')
  await runFileCases(`fields ${set}`, store, store.todo, 'FFIELDS_ONLY', (c) => runCase(c, files.get(c.id)!), save, t0)
}
