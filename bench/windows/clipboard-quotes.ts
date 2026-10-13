/** Plain text copied then pasted without Excel's rich clipboard formats, to test reversibility. */
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { writeXlsx } from '../../src/xlsx.js'
import { quoteCases } from '../coverage/clipboard/cases.js'
import { metaOf, OUT, runJob, withWindowsExcel } from './remote.js'
import { writeAnswers } from '../coverage/stored.js'
await withWindowsExcel('ru-RU', async (env) => {
  mkdirSync(OUT, { recursive: true })
  const cases = quoteCases()
  const books = await Promise.all(cases.map(async (c) => {
    const book = structuredClone(c.book)
    book.sheets[0].cells.D4 = { value: 'sentinel' }
    const file = join(OUT, `quotes-${c.id}.xlsx`)
    writeFileSync(file, (await writeXlsx({ book })).data)
    return { file, steps: [{ do: 'copyText', sheet: 'S', range: c.range, at: 'D1', read: 'D1:E4' }] }
  }))
  const got = await runJob('clipboard-quotes', books)
  const answers = Object.fromEntries(cases.map((c, i) => {
    const a = got[i]
    if (!a || !('answers' in a)) throw new Error(JSON.stringify(a))
    return [c.id, a.answers[0]]
  }))
  mkdirSync('test/excel/win/clipboard/ru-RU', { recursive: true })
  writeAnswers('test/excel/win/clipboard/ru-RU/quotes-held.json', { meta: metaOf(env, 'ru-RU'), cases, answers })
  console.log(JSON.stringify(answers))
})
