/**
 * Excel's own formula text in a locale, for the locale data of `src/locale/` and its tests:
 *
 *   npx tsx bench/locale/snapshot.ts ru     Excel with the Russian interface and locale → docs/locale/excel-ru-RU.json
 *
 * Under the Excel lock, Excel is started in the locale (`ensureExcel`), its international settings are
 * read, and the probe books (`probes.ts`) are opened one part at a time: column A of the case sheet is read
 * as `formula` (English) and `formula local`, nothing is calculated, the book is closed without saving.
 * Before and after each part Excel is checked for windows that are not workbooks; one there stops the run
 * (nothing is clicked), except Excel's "problem with some content" window when it rejects a book: it is
 * closed (No) and the part is split down to the formulas Excel does not load, which are listed in `rejected`.
 * A formula Excel deletes while loading (CALL) leaves an empty cell and is listed in `dropped`.
 * Then `npx tsx bench/locale/generate.ts ru` writes the locale file. `LOCALE_PARTS=2` runs only the first
 * parts, into `bench/coverage/out/` (a trial).
 */
import { execFileSync } from 'node:child_process'
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { writeBook } from '../coverage/book.js'
import { ensureExcel, excelDialogs, excelVersion, runScriptWatched, withExcelLock, type Locale } from '../coverage/excel.js'
import type { Probe } from '../coverage/probes.js'
import { serialize, snapshotFile, type Snapshot } from './derive.js'
import { buildProbes, SHEETS } from './probes.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const OUT = join(HERE, '../coverage/out')
const TAGS: Record<Locale, string> = { en: 'en-US', ru: 'ru-RU' }

const stripEq = (f: unknown) => (typeof f === 'string' && f.startsWith('=') ? f.slice(1) : null)

function windows(when: string): void {
  const d = excelDialogs()
  if (d.length) throw new Error(`${when}: Excel has a window that is not a workbook, stopped: ${d.map((x) => `${x.subrole} «${x.text}»`).join('; ')}`)
}

async function snapshot(locale: Locale): Promise<void> {
  const tag = TAGS[locale]
  if (!tag) throw new Error('locale: en | ru')
  const { parts: all, functions } = buildProbes()
  const trial = Number(process.env.LOCALE_PARTS) || 0
  const parts = trial ? all.slice(0, trial) : all
  const t0 = Date.now()
  await ensureExcel(locale)
  windows('before the run')
  const international = JSON.parse(execFileSync('osascript', [join(HERE, 'excel-international.applescript')], { encoding: 'utf8', timeout: 60_000 })) as Snapshot['international']
  const formulas: Snapshot['formulas'] = {}
  const rejected: string[] = []
  const dropped: string[] = []
  mkdirSync(OUT, { recursive: true })
  const run = async (cases: Probe[], label: string): Promise<void> => {
    windows(`part ${label}`)
    const book = join(OUT, `locale-${locale}-${label}.xlsx`)
    // the script reads a column of two cells at least
    const sent = cases.length === 1 ? [...cases, { id: 'pad', group: 'pad', formula: '0' }] : cases
    await writeBook(sent, book, SHEETS)
    const out = await runScriptWatched(join(HERE, 'excel-formulas.applescript'), book, sent.length)
    windows(`after part ${label}`)
    if (out !== 'rejected') {
      writeFileSync(book.replace(/\.xlsx$/, '.json'), out)
      const [en, loc] = JSON.parse(out) as [unknown[][], unknown[][]]
      cases.forEach((c, i) => {
        // a formula Excel deletes when loading the book (CALL) leaves an empty cell
        if (en[i][0] === '' && loc[i][0] === '') return void dropped.push(c.id)
        const pair = [stripEq(en[i][0]), stripEq(loc[i][0])]
        if (pair[0] === null || pair[1] === null) throw new Error(`${c.id} ${c.formula}: Excel gives ${JSON.stringify([en[i][0], loc[i][0]])}`)
        formulas[c.id] = pair as [string, string]
      })
      return
    }
    if (cases.length === 1) {
      rejected.push(cases[0].id)
      return void console.log(`Excel does not load ${cases[0].id}: ${cases[0].formula}`)
    }
    const mid = cases.length >> 1
    await run(cases.slice(0, mid), `${label}a`)
    await run(cases.slice(mid), `${label}b`)
  }
  for (let k = 0; k < parts.length; k++) {
    await run(parts[k], String(k + 1))
    console.log(`part ${k + 1}/${parts.length}: ${parts[k].length} formulas, ${((Date.now() - t0) / 1000).toFixed(1)} s`)
  }
  const snap: Snapshot = {
    meta: { excel: excelVersion(), mode: locale === 'en' ? 'en-US (-AppleLanguages (en) -AppleLocale en_US)' : 'ru-RU (the system language and locale)', date: new Date().toISOString().slice(0, 10), seconds: Math.round((Date.now() - t0) / 100) / 10 },
    international,
    functions,
    formulas,
    rejected,
    dropped,
  }
  const file = trial ? join(OUT, `excel-${tag}.trial.json`) : fileURLToPath(snapshotFile(tag))
  mkdirSync(join(file, '..'), { recursive: true })
  writeFileSync(file, serialize(snap))
  console.log(`${Object.keys(formulas).length} formulas, ${rejected.length} rejected, ${dropped.length} dropped → ${file}`)
}

await withExcelLock(() => snapshot((process.argv[2] ?? 'ru') as Locale))
