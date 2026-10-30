/** The overlapping `<col>` cases in Excel for Mac en-US, under the shared lock: `npx tsx bench/coverage/cols/mac.ts`. */
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { ensureExcel, excelVersion, runWatched, settleRejected, withExcelLock } from '../excel.js'
import { writeAnswers } from '../stored.js'
import { partsHash } from '../xlsx/sent.js'
import { CASES, caseFile, type ColRead, type ColsAnswer, letters, savedCols } from './cases.js'

await withExcelLock(async () => {
  await ensureExcel('en-US')
  const answers: Record<string, ColsAnswer> = {}
  const sent: Record<string, string> = {}
  for (const c of CASES) {
    const data = await caseFile(c)
    sent[c.id] = await partsHash(data)
    answers[c.id] = await runWatched<ColsAnswer>({
      script: new URL('./cols.applescript', import.meta.url).pathname,
      label: c.id,
      folder: 'fcols',
      empty: true,
      book: (file) => writeFileSync(file, data),
      args: (dir) => [join(dir, 'saved.xlsx'), 'saved.xlsx', JSON.stringify(letters(c))],
      timeout: 60_000,
      poll: 500,
      fatal: true,
      after: 'case',
      repair: async (stop, text) => {
        await stop()
        await settleRejected()
        return { repaired: text }
      },
      done: async (out, dir) => {
        const got = JSON.parse(out) as ColRead[]
        return { read: Object.fromEntries(letters(c).map((l, i) => [l, got[i]])), saved: await savedCols(new Uint8Array(readFileSync(join(dir, 'saved.xlsx')))) }
      },
    })
    console.log(c.id, JSON.stringify(answers[c.id]))
  }
  const dir = new URL('../../../test/excel/cols/', import.meta.url).pathname
  mkdirSync(dir, { recursive: true })
  writeAnswers(`${dir}overlap.json`, { meta: { version: excelVersion(), platform: 'Mac', locale: 'en-US' }, sent, answers })
})
