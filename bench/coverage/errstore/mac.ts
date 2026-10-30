/** The formulas of `cases.ts` in Excel for Mac en-US, under the shared lock: `npx tsx bench/coverage/errstore/mac.ts`. */
import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { openArchive } from '../../../src/xlsx/zip/read.js'
import { ensureExcel, excelVersion, runWatched, withExcelLock } from '../excel.js'
import { writeAnswers } from '../stored.js'
import { partsHash } from '../xlsx/sent.js'
import { AFTER, BEFORE, bookFile, type ErrStoreAnswer } from './cases.js'

await withExcelLock(async () => {
  await ensureExcel('en-US')
  const data = await bookFile()
  const answer = await runWatched<ErrStoreAnswer>({
    script: new URL('./errstore.applescript', import.meta.url).pathname,
    label: 'errstore',
    folder: 'ferrs',
    empty: true,
    book: (file) => writeFileSync(file, data),
    args: (dir) => [join(dir, 'saved.xlsx'), 'saved.xlsx', JSON.stringify([BEFORE, AFTER])],
    timeout: 60_000,
    poll: 500,
    fatal: true,
    after: 'script',
    repair: async (stop) => {
      await stop()
      throw new Error('Excel would repair the book')
    },
    done: async (out, dir) => {
      const sheet = (await openArchive(new Uint8Array(readFileSync(join(dir, 'saved.xlsx')))).text('xl/worksheets/sheet1.xml'))!
      return { ...(JSON.parse(out) as Omit<ErrStoreAnswer, 'saved'>), saved: [...sheet.matchAll(/<c r="[^"]*"[^>]*>.*?<\/c>/g)].map((m) => m[0]) }
    },
  })
  writeAnswers(new URL('../../../test/excel/errstore/mac.json', import.meta.url).pathname, { meta: { version: excelVersion(), platform: 'Mac', locale: 'en-US' }, sent: await partsHash(data), ...answer })
  console.log(JSON.stringify(answer, null, 1))
})
