import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { attrsAt } from '../../../src/xlsx/xml/scan.js'
import { openArchive } from '../../../src/xlsx/zip/read.js'
import { metaOf, OUT, runJob, withWindowsExcel, type Step } from '../../windows/remote.js'
import { ensureExcel, excelVersion, runSavingScript, withExcelLock } from '../excel.js'
import { ROOT } from '../paths.js'
import { writeAnswers } from '../stored.js'
import { DEFAULT_WIDTH_CASES, defaultWidthFile } from './default-column-width.js'

const LOCALE = 'ru-RU'

async function stored(data: Uint8Array) {
  const xml = await openArchive(data).text('xl/worksheets/sheet1.xml')
  return attrsAt(xml ?? '', 'sheetFormatPr')
}

if (process.argv[2] === 'mac') {
  await withExcelLock(async () => {
    await ensureExcel(LOCALE)
    const script = fileURLToPath(new URL('./excel-default-column-width.applescript', import.meta.url))
    const entries = []
    for (const c of DEFAULT_WIDTH_CASES) {
      const r = await runSavingScript(`default-width-${c.id}`, script, await defaultWidthFile(c), [])
      entries.push([c.id, 'rejected' in r ? r : { columns: JSON.parse(r.out), saved: await stored(r.saved) }])
    }
    writeAnswers(join(ROOT, 'test/excel/xlsx/default-column-width.json'), { meta: { version: excelVersion(), locale: LOCALE }, answers: Object.fromEntries(entries) })
  })
} else if (process.argv[2] === 'windows') {
  await withWindowsExcel(LOCALE, async (env) => {
    mkdirSync(OUT, { recursive: true })
    const steps: Step[] = [{ do: 'widths', sheet: 'S', columns: ['A', 'B'] }, { do: 'shown', sheet: 'S', cells: ['A1', 'B1'] }, { do: 'saveAs' }]
    const books = await Promise.all(DEFAULT_WIDTH_CASES.map(async (c) => {
      const file = join(OUT, `default-width-${c.id}.xlsx`)
      writeFileSync(file, await defaultWidthFile(c))
      return { file, steps }
    }))
    const results = await runJob('xlsx-default-column-width', books)
    const entries = await Promise.all(DEFAULT_WIDTH_CASES.map(async ({ id }, index) => {
      const r = results[index]
      if (!r) throw new Error(`${id}: Excel did not finish this case`)
      if ('rejected' in r) return [id, { rejected: r.rejected }]
      if (!r.saved) throw new Error(`${id}: Excel did not save the book`)
      return [id, { widths: r.answers[0], shown: r.answers[1], saved: await stored(r.saved) }]
    }))
    writeAnswers(join(ROOT, 'test/excel/win/xlsx/default-column-width.json'), { meta: metaOf(env, LOCALE), answers: Object.fromEntries(entries) })
  })
} else throw new Error('Specify mac or windows')
