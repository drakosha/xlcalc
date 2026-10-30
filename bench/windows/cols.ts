/** The overlapping `<col>` cases of `bench/coverage/cols` in Excel for Windows en-US: how each opened, its columns, the file saved. */
import { writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { CASES, caseFile, type ColsAnswer, letters, savedCols } from '../coverage/cols/cases.js'
import { writeAnswers } from '../coverage/stored.js'
import { partsHash } from '../coverage/xlsx/sent.js'
import { metaOf, OUT, runJob, withWindowsExcel } from './remote.js'

await withWindowsExcel('en-US', async (env) => {
  const sent: Record<string, string> = {}
  const books = await Promise.all(
    CASES.map(async (c, i) => {
      const data = await caseFile(c)
      sent[c.id] = await partsHash(data)
      const file = join(OUT, `cols-${i}.xlsx`)
      writeFileSync(file, data)
      return { file, steps: [{ do: 'opened' }, { do: 'widths', sheet: 'S', columns: letters(c) }, { do: 'cells', sheet: 'S', cells: letters(c).map((l) => `${l}1`) }, { do: 'saveAs' }] }
    }),
  )
  const got = await runJob('cols', books)
  const answers: Record<string, ColsAnswer & { opened?: unknown }> = {}
  for (const [i, c] of CASES.entries()) {
    const a = got[i]
    if (!a || !('answers' in a) || !a.saved) {
      answers[c.id] = { repaired: a && 'rejected' in a ? a.rejected : 'no answer' }
      continue
    }
    const [opened, widths, cells] = a.answers as [unknown, [number, number, boolean][], unknown[][]]
    answers[c.id] = { opened, read: Object.fromEntries(letters(c).map((l, k) => [l, [widths[k][0], String(cells[k][1])]])), saved: await savedCols(a.saved) }
    console.log(c.id, JSON.stringify(answers[c.id]))
  }
  writeAnswers('test/excel/win/cols/overlap.json', { meta: metaOf(env, 'en-US'), sent, answers })
})
