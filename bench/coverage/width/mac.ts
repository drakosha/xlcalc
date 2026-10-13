/** Fixed widths read through Excel for Mac's object model under the shared lock. */
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { writeXlsx } from '../../../src/xlsx.js'
import { COLUMNS, SET_WIDTHS, widthBook } from './cases.js'
import { savedColumns } from '../fit.js'
import { ensureExcel, excelVersion, runWatched, withExcelLock } from '../excel.js'
import { writeAnswers } from '../stored.js'
await withExcelLock(async () => {
  await ensureExcel('ru-RU')
  const answer = await runWatched({
    script: new URL('./width.applescript', import.meta.url).pathname,
    label: 'column widths', folder: 'fwidth', empty: true,
    book: async (file) => { writeFileSync(file, (await writeXlsx({ book: widthBook() })).data) },
    args: (dir) => [join(dir, 'saved.xlsx'), 'saved.xlsx', JSON.stringify(SET_WIDTHS.map((w, i) => [COLUMNS[i + 11], w]))],
    timeout: 30_000, poll: 500, fatal: true,
    repair: async (stop) => { await stop(); throw new Error('Excel refused the widths book') },
    done: async (out, dir) => ({ ...JSON.parse(out) as object, saved: Object.fromEntries(await savedColumns(new Uint8Array(readFileSync(join(dir, 'saved.xlsx'))))) }),
  })
  const dir = new URL('../../../test/excel/width/', import.meta.url).pathname
  mkdirSync(dir, { recursive: true })
  writeAnswers(`${dir}width-held.json`, { meta: { version: excelVersion(), platform: 'Mac', locale: 'ru-RU' }, book: widthBook(), set: SET_WIDTHS, ...answer })
  console.log(JSON.stringify(answer))
})
