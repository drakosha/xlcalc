/**
 * Excel for Windows (the machine `windows-host`) computes the Mac harness's cases; answers go to `test/excel/win/`.
 *
 *   npx tsx bench/windows/run.ts env                   the machine's settings and Excel's, as it starts now
 *   npx tsx bench/windows/run.ts measurement en|ru     the coverage measurement into coverage/excel-<en|ru>.json
 *   npx tsx bench/windows/run.ts probes                the probes (en-US) into probes.json
 *   npx tsx bench/windows/run.ts set <name> [ru]       a case set into cases/[ru-RU/]<name>.json
 *   npx tsx bench/windows/run.ts legacy <name>         a legacy set (en-US) into legacy/<name>.json
 *   npx tsx bench/windows/run.ts functions en|ru       all of the above that the Mac has answers for in that locale
 *
 * Each command holds the machine's lock and sets its locale first (restarting it when the system locale changes).
 * FCOV_FRESH=1 runs every case again; XLWIN_HOST names another machine.
 */
import { readdirSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { stepFunctions } from './functions.js'
import { env, withWindowsExcel, type Tag } from './remote.js'

const ROOT = join(fileURLToPath(new URL('.', import.meta.url)), '../..')
const tagOf = (x: string | undefined): Tag => (x === 'ru' ? 'ru-RU' : 'en-US')

/** The sets with Mac answers in a locale, by their files in `test/excel/`. */
const macSets = (dir: string) => readdirSync(join(ROOT, 'test/excel', dir)).filter((f) => f.endsWith('.json')).map((f) => f.slice(0, -5))

/** The formula targets with Mac answers in a locale. */
const functionTargets = (tag: Tag) => [
  'measurement',
  ...(tag === 'en-US' ? ['probes', ...macSets('legacy').filter((s) => s !== 'measurement').map((s) => `legacy/${s}`)] : []),
  ...macSets(tag === 'en-US' ? 'cases' : 'cases/ru-RU'),
]

const [step, a, b] = process.argv.slice(2)
const t0 = Date.now()
if (step === 'env') console.log(env())
else if (step === 'measurement') await withWindowsExcel(tagOf(a), (e) => stepFunctions(['measurement'], tagOf(a), e))
else if (step === 'probes') await withWindowsExcel('en-US', (e) => stepFunctions(['probes'], 'en-US', e))
else if (step === 'set' && a) await withWindowsExcel(tagOf(b), (e) => stepFunctions([a], tagOf(b), e))
else if (step === 'legacy' && a) await withWindowsExcel('en-US', (e) => stepFunctions([`legacy/${a}`], 'en-US', e))
else if (step === 'functions') await withWindowsExcel(tagOf(a), (e) => stepFunctions(functionTargets(tagOf(a)), tagOf(a), e))
else {
  console.error('step: env | measurement en|ru | probes | set <name> [ru] | legacy <name> | functions en|ru')
  process.exit(2)
}
console.log(`done in ${((Date.now() - t0) / 1000).toFixed(1)} s`)
