import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { attrsAt } from '../../../src/xlsx/xml/scan.js'
import { openArchive } from '../../../src/xlsx/zip/read.js'
import { metaOf, OUT, runJob, withWindowsExcel, type Step } from '../../windows/remote.js'
import { ensureExcel, excelVersion, runSavingScript, withExcelLock } from '../excel.js'
import { ROOT } from '../paths.js'
import { loadStored, writeAnswers } from '../stored.js'
import { DEFAULT_WIDTH_CASES, defaultWidthFile } from './default-column-width.js'

const LOCALE = 'ru-RU'

async function stored(data: Uint8Array) {
  const xml = await openArchive(data).text('xl/worksheets/sheet1.xml')
  return attrsAt(xml ?? '', 'sheetFormatPr')
}

function pending(file: string) {
  const previous = loadStored<{ answers: Record<string, unknown> }>(file)
  const answers = previous?.answers ?? {}
  const cases = DEFAULT_WIDTH_CASES.filter(({ id }) => !(id in answers))
  return {
    cases,
    save(meta: unknown, entries: [string, unknown][]) {
      writeAnswers(file, { meta, answers: { ...answers, ...Object.fromEntries(entries) } })
    },
  }
}

if (process.argv[2] === 'mac') {
  await withExcelLock(async () => {
    const { cases, save } = pending(join(ROOT, 'test/excel/xlsx/default-column-width.json'))
    if (!cases.length) return
    await ensureExcel(LOCALE)
    const script = fileURLToPath(new URL('./excel-default-column-width.applescript', import.meta.url))
    const entries: [string, unknown][] = []
    for (const c of cases) {
      const r = await runSavingScript(`default-width-${c.id}`, script, await defaultWidthFile(c), [])
      entries.push([c.id, 'rejected' in r ? r : { columns: JSON.parse(r.out), saved: await stored(r.saved) }])
    }
    save({ version: excelVersion(), locale: LOCALE }, entries)
  })
} else if (process.argv[2] === 'windows') {
  await withWindowsExcel(LOCALE, async (env) => {
    const { cases, save } = pending(join(ROOT, 'test/excel/win/xlsx/default-column-width.json'))
    if (!cases.length) return
    mkdirSync(OUT, { recursive: true })
    const steps: Step[] = [{ do: 'widths', sheet: 'S', columns: ['A', 'B'] }, { do: 'shown', sheet: 'S', cells: ['A1', 'B1'] }, { do: 'saveAs' }]
    const books = await Promise.all(cases.map(async (c) => {
      const file = join(OUT, `default-width-${c.id}.xlsx`)
      writeFileSync(file, await defaultWidthFile(c))
      return { file, steps }
    }))
    const results = await runJob('xlsx-default-column-width', books)
    const entries = await Promise.all(cases.map(async ({ id }, index): Promise<[string, unknown]> => {
      const r = results[index]
      if (!r) throw new Error(`${id}: Excel did not finish this case`)
      if ('rejected' in r) return [id, { rejected: r.rejected }]
      if (!r.saved) throw new Error(`${id}: Excel did not save the book`)
      return [id, { widths: r.answers[0], shown: r.answers[1], saved: await stored(r.saved) }]
    }))
    save(metaOf(env, LOCALE), entries)
  })
} else throw new Error('Specify mac or windows')
