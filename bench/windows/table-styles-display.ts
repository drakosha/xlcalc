import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { writeAnswers } from '../coverage/stored.js'
import { DISPLAY_CASES, displayFile } from '../coverage/table-styles/display-held.js'
import { winFile } from './functions.js'
import { metaOf, OUT, runJob, withWindowsExcel, type Step } from './remote.js'

await withWindowsExcel('ru-RU', async (env) => {
  mkdirSync(OUT, { recursive: true })
  const books = await Promise.all(DISPLAY_CASES.map(async ({ id, sheet, range }) => {
    const file = join(OUT, `table-style-display-${id}.xlsx`)
    writeFileSync(file, await displayFile(id))
    const steps: Step[] = [{ do: 'full' }, { do: 'cond', op: 'looks', read: [[sheet, range]], borders: true }]
    return { file, steps }
  }))
  const results = await runJob('table-style-display', books)
  const entries = DISPLAY_CASES.map(({ id }, i) => {
    const result = results[i]
    if (!result || 'rejected' in result) throw new Error(`${id}: Excel did not open the synthetic workbook`)
    return [id, result.answers[1]]
  })
  writeAnswers(winFile('xlsx/table-style-display-held.json'), { meta: metaOf(env, 'ru-RU'), answers: Object.fromEntries(entries) })
})
