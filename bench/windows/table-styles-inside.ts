import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { writeAnswers } from '../coverage/stored.js'
import { INSIDE_CASES, insideFile } from '../coverage/table-styles/inside-held.js'
import { winFile } from './functions.js'
import { metaOf, OUT, runJob, withWindowsExcel, type Step } from './remote.js'

// Inside borders, patterns and number formats of custom table styles, as the display format shows them.
const set = 'table-style-inside-held'
await withWindowsExcel('ru-RU', async (env) => {
  mkdirSync(OUT, { recursive: true })
  const books = await Promise.all(INSIDE_CASES.map(async ({ id, sheet, range }) => {
    const file = join(OUT, `table-style-${id}.xlsx`)
    writeFileSync(file, await insideFile())
    const steps: Step[] = [{ do: 'full' }, { do: 'cond', kind: 'looks', read: [[sheet, range]], borders: true, patterns: true }, { do: 'saveAs' }]
    return { file, steps }
  }))
  const results = await runJob(set, books)
  const entries = INSIDE_CASES.map(({ id }, i) => {
    const result = results[i]
    if (!result) throw new Error(`${id}: Excel did not finish the case`)
    if ('rejected' in result) throw new Error(`${id}: ${result.rejected}`)
    return [id, result.answers[1]]
  })
  writeAnswers(winFile(`xlsx/${set}.json`), { meta: metaOf(env, 'ru-RU'), answers: Object.fromEntries(entries) })
})
