/**
 * Fields of a workbook in Excel for Windows (`bench/coverage/fields/`): each case's file opened, each sheet's tab colour
 * (`Tab.Color`, none where `Tab.ColorIndex` is `xlColorIndexNone`), gridlines and headings (the active window's
 * `DisplayGridlines` and `DisplayHeadings` with the sheet active) read, and the book saved. The answers go to
 * `test/excel/win/fields/<set>.json` and the files Excel saved beside them, `<set>/<id>.xlsx`.
 */
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fieldsFile, fieldsFileOf, savedDir, type FieldsAnswer, type FieldsCase, type FieldsSet, type SheetFields } from '../coverage/fields/cases.js'
import { fromBgr } from '../coverage/cond/cases.js'
import { answeredLine, answerSaver, keptSame, loadStored, setModule } from '../coverage/stored.js'
import { inJobs, metaOf, savedInto, type WinEnv } from './remote.js'

export async function stepFields(set: string, e: WinEnv): Promise<void> {
  const cases = (await setModule<FieldsSet>('fields', set)).buildFieldsCases()
  const file = fieldsFile(set)
  const answers = keptSame<FieldsCase, FieldsAnswer>({ cases }, loadStored(file))
  const t0 = Date.now()
  const save = answerSaver(file, metaOf(e, 'en-US'), true, t0, cases, answers)
  const todo = cases.filter((c) => !answers[c.id])
  const results: [FieldsCase, unknown[], Uint8Array][] = []
  await inJobs(
    `fields-${set}`,
    todo,
    async (c, f) => (writeFileSync(f, await fieldsFileOf(c)), { file: f, steps: [{ do: 'sheetFields' }, { do: 'saveAs' }] }),
    savedInto(results, (c, why) => `fields ${set} ${c.id}: the book did not run: ${why}`),
    save,
  )
  mkdirSync(savedDir(set), { recursive: true })
  for (const [c, got, saved] of results) {
    const raw = got[0] as [string, number | null, boolean, boolean][]
    answers[c.id] = { sheets: raw.map(([name, bgr, grid, heads]): SheetFields => [name, bgr === null ? null : fromBgr(bgr), grid, heads]) }
    writeFileSync(join(savedDir(set), `${c.id}.xlsx`), saved)
  }
  save()
  console.log(answeredLine(`fields ${set}`, todo.length, answers, cases))
}
