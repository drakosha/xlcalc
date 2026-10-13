/** Column widths from COM in points, without a screen measurement. */
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { writeXlsx } from '../../src/xlsx.js'
import { COLUMNS, SET_WIDTHS, widthBook } from '../coverage/width/cases.js'
import { savedColumns } from '../coverage/fit.js'
import { writeAnswers } from '../coverage/stored.js'
import { metaOf, OUT, runJob, withWindowsExcel } from './remote.js'
await withWindowsExcel('ru-RU', async (env) => {
  mkdirSync(OUT, { recursive: true })
  const file = join(OUT, 'width.xlsx')
  writeFileSync(file, (await writeXlsx({ book: widthBook() })).data)
  const steps = [
    { do: 'standardWidth', sheet: 'S' },
    ...SET_WIDTHS.map((width, i) => ({ do: 'setColumnWidth', sheet: 'S', column: `${COLUMNS[i + 11]}:${COLUMNS[i + 11]}`, width })),
    { do: 'widths', sheet: 'S', columns: COLUMNS }, { do: 'saveAs' },
  ]
  const [answer] = await runJob('width-held', [{ file, steps }])
  if (!answer || !('answers' in answer) || !answer.saved) throw new Error(JSON.stringify(answer))
  const saved = Object.fromEntries(await savedColumns(answer.saved))
  mkdirSync('test/excel/win/width', { recursive: true })
  writeAnswers('test/excel/win/width/width-held.json', { meta: metaOf(env, 'ru-RU'), book: widthBook(), set: SET_WIDTHS, standardWidth: answer.answers[0], columns: answer.answers.at(-2), saved })
  console.log(JSON.stringify(answer.answers))
})
