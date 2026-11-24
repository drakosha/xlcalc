/** Open the raw format codes in Excel and read the displayed cells and their codes where it opens them. */
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { FORMAT_CONDITION_CASES, conditionFormatFile } from '../coverage/xlsx/number-format-conditions.js'
import { writeAnswers } from '../coverage/stored.js'
import { winFile } from './functions.js'
import { metaOf, OUT, runJob, withWindowsExcel, type Step } from './remote.js'

await withWindowsExcel('ru-RU', async (env) => {
  mkdirSync(OUT, { recursive: true })
  const steps: Step[] = [{ do: 'shown', sheet: 'S', cells: ['A1'], codes: true }, { do: 'saveAs' }]
  const books = await Promise.all(FORMAT_CONDITION_CASES.map(async ({ id, code }) => {
    const file = join(OUT, `number-format-conditions-${id}.xlsx`)
    writeFileSync(file, await conditionFormatFile(code))
    return { file, steps }
  }))
  const results = await runJob('xlsx-number-format-conditions', books)
  const entries = FORMAT_CONDITION_CASES.map(({ id }, index) => {
    const r = results[index]
    if (!r) throw new Error(`${id}: Excel did not finish this case`)
    return [id, 'rejected' in r ? { rejected: r.rejected } : { shown: r.answers[0], saved: !!r.saved }]
  })
  writeAnswers(winFile('xlsx/number-format-conditions.json'), { meta: metaOf(env, 'ru-RU'), answers: Object.fromEntries(entries) })
})
