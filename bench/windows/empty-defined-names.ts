/** Read Excel's names and formula value on opening, then the defined names it saves. */
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { attributes, within } from '../../src/xlsx/xml/scan.js'
import { openArchive } from '../../src/xlsx/zip/read.js'
import { writeAnswers } from '../coverage/stored.js'
import { EMPTY_NAME_CASES, emptyNameFile } from '../coverage/xlsx/empty-defined-names.js'
import { winFile } from './functions.js'
import { metaOf, OUT, runJob, withWindowsExcel, type Step } from './remote.js'

await withWindowsExcel('ru-RU', async (env) => {
  mkdirSync(OUT, { recursive: true })
  const steps: Step[] = [
    { do: 'names' },
    { do: 'full' },
    { do: 'values', sheet: 'S', range: 'B1' },
    { do: 'saveAs' },
  ]
  const files = await Promise.all(EMPTY_NAME_CASES.map(async (c) => {
    const file = join(OUT, `empty-defined-names-${c.id}.xlsx`)
    writeFileSync(file, await emptyNameFile(c))
    return { file, steps }
  }))
  const results = await runJob('xlsx-empty-defined-names', files)
  const entries = await Promise.all(EMPTY_NAME_CASES.map(async ({ id }, index) => {
    const result = results[index]
    if (!result) throw new Error(`${id}: the Excel job did not finish`)
    if ('rejected' in result) return [id, { rejected: result.rejected }] as const
    if (!result.saved) throw new Error(`${id}: Excel did not save the workbook`)
    const xml = (await openArchive(result.saved).text('xl/workbook.xml'))!
    const saved = within(xml, 'definedNames', 'definedName').map((e) => ({ ...attributes(e.attrs), ref: e.body ?? '' }))
    const [names, , value] = result.answers
    return [id, { names, value, saved }] as const
  }))
  writeAnswers(winFile('xlsx/empty-defined-names.json'), { meta: metaOf(env, 'ru-RU'), answers: Object.fromEntries(entries) })
})
