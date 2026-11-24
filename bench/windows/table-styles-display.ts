import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { writeAnswers } from '../coverage/stored.js'
import { DISPLAY_CASES, displayFile } from '../coverage/table-styles/display-held.js'
import { REGION_CASES, regionFile } from '../coverage/table-styles/regions-held.js'
import { winFile } from './functions.js'
import { metaOf, OUT, runJob, withWindowsExcel, type Step } from './remote.js'

const regions = process.argv[2] === 'regions'
const cases = regions ? REGION_CASES.map((c) => ({ ...c, file: () => regionFile(c.id) })) : DISPLAY_CASES.map((c) => ({ ...c, file: () => displayFile(c.id) }))
const set = regions ? 'table-style-regions-held' : 'table-style-display-held'

await withWindowsExcel('ru-RU', async (env) => {
  mkdirSync(OUT, { recursive: true })
  const books = await Promise.all(cases.map(async ({ id, sheet, range, file: input }) => {
    const file = join(OUT, `table-style-display-${id}.xlsx`)
    writeFileSync(file, await input())
    const steps: Step[] = [{ do: 'full' }, { do: 'cond', kind: 'looks', read: [[sheet, range]], borders: true }]
    return { file, steps }
  }))
  const results = await runJob(set, books)
  const entries = cases.map(({ id }, i) => {
    const result = results[i]
    if (!result) throw new Error(`${id}: Excel did not finish the case`)
    if ('rejected' in result) throw new Error(`${id}: ${result.rejected}`)
    return [id, result.answers[1]]
  })
  writeAnswers(winFile(`xlsx/${set}.json`), { meta: metaOf(env, 'ru-RU'), answers: Object.fromEntries(entries) })
})
