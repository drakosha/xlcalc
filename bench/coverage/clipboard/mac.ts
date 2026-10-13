/** Clipboard reads through the Mac object model; no activate, selection or keyboard calls. */
import { mkdirSync, writeFileSync } from 'node:fs'
import { writeXlsx } from '../../../src/xlsx.js'
import { ensureExcel, excelVersion, runWatched, withExcelLock } from '../excel.js'
import { clipboardCases } from './cases.js'
import { writeAnswers } from '../stored.js'
await withExcelLock(async () => {
  try {
    for (const tag of ['en-US', 'ru-RU']) {
      await ensureExcel(tag)
      const cases = clipboardCases(tag)
      const answers: Record<string, { text: string }> = {}
      for (const c of cases) {
        const text = await runWatched<string>({
          script: new URL('./copy.applescript', import.meta.url).pathname,
          label: `clipboard ${c.id}`, folder: 'fcopy', empty: true,
          book: async (f) => { writeFileSync(f, (await writeXlsx({ book: c.book })).data) },
          args: () => [c.range, JSON.stringify((c.steps ?? []).map((s) => [s.do, s.range]))], timeout: 20_000, poll: 500, fatal: true,
          repair: async (stop) => { await stop(); throw new Error('Excel refused the clipboard book') },
          done: (out) => Buffer.from(JSON.parse(out)[0] as string, 'base64').toString('utf8'),
        })
        if (text === 'xlcalc-copy-probe') throw new Error('copy range left the clipboard unchanged; no screen fallback permitted')
        answers[c.id] = { text }
      }
      const dir = new URL('../../../test/excel/clipboard/', import.meta.url).pathname
      mkdirSync(dir, { recursive: true })
      writeAnswers(`${dir}copy-${tag}.json`, { meta: { version: excelVersion(), platform: 'Mac', locale: tag }, cases, answers })
      console.log(`${tag}: ${cases.length} clipboard cases recorded`)
    }
  } finally { await ensureExcel('ru-RU') }
})
