/**
 * Excel for Windows (the machine `windows-host`) computes the Mac harness's cases; answers go to `test/excel/win/`.
 *
 *   npx tsx bench/windows/run.ts env                   the machine's settings and Excel's, as it starts now
 *   npx tsx bench/windows/run.ts measurement en|ru     the coverage measurement into coverage/excel-<en|ru>.json
 *   npx tsx bench/windows/run.ts probes                the probes (en-US) into probes.json
 *   npx tsx bench/windows/run.ts set <name> [ru]       a case set into cases/[ru-RU/]<name>.json
 *   npx tsx bench/windows/run.ts legacy <name>         a legacy set (en-US) into legacy/<name>.json
 *   npx tsx bench/windows/run.ts display <name> [ru]   what cells show into display/[ru-RU/]<name>.json
 *   npx tsx bench/windows/run.ts spill|workbook|struct <name>   a book set (en-US) into <kind>/<name>.json
 *   npx tsx bench/windows/run.ts all en|ru             everything the Mac has answers for in that locale
 *
 * Each command holds the machine's lock and sets its locale first (restarting it when the system locale changes).
 * FCOV_FRESH=1 runs every case again; XLWIN_HOST names another machine.
 */
import { readdirSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { stepDisplay, stepSpill, stepStruct, stepWorkbook } from './books.js'
import { stepFunctions } from './functions.js'
import { env, withWindowsExcel, type Tag, type WinEnv } from './remote.js'

const ROOT = join(fileURLToPath(new URL('.', import.meta.url)), '../..')
const tagOf = (x: string | undefined): Tag => (x === 'ru' ? 'ru-RU' : 'en-US')

/** The sets with Mac answers, by their files in `test/excel/<dir>`. */
const macSets = (dir: string) => readdirSync(join(ROOT, 'test/excel', dir)).filter((f) => f.endsWith('.json')).map((f) => f.slice(0, -5))

const BOOK_STEPS = { spill: stepSpill, workbook: stepWorkbook, struct: stepStruct } as const

/** Everything with Mac answers in a locale: formula cases, what cells show, and in en-US the book sets. */
async function all(tag: Tag, e: WinEnv) {
  const en = tag === 'en-US'
  const legacy = macSets('legacy').filter((s) => s !== 'measurement').map((s) => `legacy/${s}`)
  await stepFunctions(['measurement', ...(en ? ['probes', ...legacy] : []), ...macSets(en ? 'cases' : 'cases/ru-RU')], tag, e)
  for (const s of macSets(en ? 'display' : 'display/ru-RU')) await stepDisplay(s, tag, e)
  if (!en) return
  // the order runs file keeps every order Excel gave: it is no set
  for (const s of macSets('spill').filter((s) => s !== 'order-runs')) await stepSpill(s, tag, e)
  for (const s of macSets('workbook')) await stepWorkbook(s, tag, e)
  for (const s of macSets('struct')) await stepStruct(s, tag, e)
}

const [step, a, b] = process.argv.slice(2)
const t0 = Date.now()
if (step === 'env') console.log(env())
else if (step === 'measurement') await withWindowsExcel(tagOf(a), (e) => stepFunctions(['measurement'], tagOf(a), e))
else if (step === 'probes') await withWindowsExcel('en-US', (e) => stepFunctions(['probes'], 'en-US', e))
else if (step === 'set' && a) await withWindowsExcel(tagOf(b), (e) => stepFunctions([a], tagOf(b), e))
else if (step === 'legacy' && a) await withWindowsExcel('en-US', (e) => stepFunctions([`legacy/${a}`], 'en-US', e))
else if (step === 'display' && a) await withWindowsExcel(tagOf(b), (e) => stepDisplay(a, tagOf(b), e))
else if (step in BOOK_STEPS && a) await withWindowsExcel('en-US', (e) => BOOK_STEPS[step as keyof typeof BOOK_STEPS](a, 'en-US', e))
else if (step === 'all') await withWindowsExcel(tagOf(a), (e) => all(tagOf(a), e))
else {
  console.error('step: env | measurement en|ru | probes | set <name> [ru] | legacy <name> | display <name> [ru] | spill|workbook|struct <name> | all en|ru')
  process.exit(2)
}
console.log(`done in ${((Date.now() - t0) / 1000).toFixed(1)} s`)
