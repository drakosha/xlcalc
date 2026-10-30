/** The formulas of `bench/coverage/errstore` in Excel for Windows en-US: as the book opens, recalculated, after the deletion, fully calculated, and the file saved. */
import { writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { bookFile } from '../coverage/errstore/cases.js'
import { writeAnswers } from '../coverage/stored.js'
import { partsHash } from '../coverage/xlsx/sent.js'
import { openArchive } from '../../src/xlsx/zip/read.js'
import { metaOf, OUT, runJob, withWindowsExcel } from './remote.js'

await withWindowsExcel('en-US', async (env) => {
  const data = await bookFile()
  const file = join(OUT, 'errstore.xlsx')
  writeFileSync(file, data)
  const read = (range: string) => [{ do: 'formulas', sheet: 'S', range }, { do: 'values', sheet: 'S', range }]
  const steps = [
    { do: 'opened' }, ...read('C1:C12'),
    { do: 'calc' }, ...read('C1:C12'),
    { do: 'struct', op: 'deleteColumns', sheet: 'S', range: 'A1' }, { do: 'struct', op: 'deleteSheet', sheet: 'T' }, ...read('B1:B12'),
    { do: 'saveAs' },
    { do: 'full' }, ...read('B1:B12'),
  ]
  const [a] = await runJob('errstore', [{ file, steps }])
  if (!a || !('answers' in a) || !a.saved) throw new Error(JSON.stringify(a))
  const sheet = (await openArchive(a.saved).text('xl/worksheets/sheet1.xml'))!
  const [opened, f1, v1, , f2, v2, , , f3, v3, , , f4, v4] = a.answers
  const answer = { opened, opening: [f1, v1], recalculated: [f2, v2], deleted: [f3, v3], full: [f4, v4], saved: [...sheet.matchAll(/<c r="[^"]*"[^>]*>.*?<\/c>/g)].map((m) => m[0]) }
  writeAnswers('test/excel/win/errstore/win.json', { meta: metaOf(env, 'en-US'), sent: await partsHash(data), ...answer })
  console.log(JSON.stringify(answer))
})
