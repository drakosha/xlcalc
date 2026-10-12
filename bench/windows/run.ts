/**
 * Excel for Windows (the machine `windows-host`) computes the Mac harness's cases; answers go to `test/excel/win/`.
 *
 *   npx tsx bench/windows/run.ts env                   the machine's settings and Excel's, as it starts now
 *   npx tsx bench/windows/run.ts measurement L         the coverage measurement into coverage/excel-<en|ru|tag>.json
 *   npx tsx bench/windows/run.ts probes                the probes (en-US) into probes.json
 *   npx tsx bench/windows/run.ts set <name> [L]        a case set into cases/[<tag>/]<name>.json
 *   npx tsx bench/windows/run.ts legacy <name>         a legacy set (en-US) into legacy/<name>.json
 *   npx tsx bench/windows/run.ts display <name> [L]    what cells show into display/[<tag>/]<name>.json
 *   npx tsx bench/windows/run.ts tables <name> [L]     tables and their references into tables/[<tag>/]<name>.json
 *   npx tsx bench/windows/run.ts spill|workbook|struct <name>   a book set (en-US) into <kind>/<name>.json
 *   npx tsx bench/windows/run.ts parts <name>          notes, hyperlinks and frozen panes (en-US) into parts/<name>.json and parts/<name>/
 *   npx tsx bench/windows/run.ts hidden <name>         hidden sheets (en-US) into hidden/<name>.json and hidden/<name>/
 *   npx tsx bench/windows/run.ts links <name>          links to other workbooks, opened without updating them (en-US), into links/<name>.json and links/<name>/
 *   npx tsx bench/windows/run.ts lines <name>          formats of rows and columns (en-US) into lines/<name>.json and lines/<name>/
 *   npx tsx bench/windows/run.ts xlsx filters        the filter files computed and saved by Excel into xlsx/filters/
 *   npx tsx bench/windows/run.ts xlsx formats [L]    the codes of the built-in number formats into xlsx/[<tag>/]formats.json
 *   npx tsx bench/windows/run.ts entry <name> [L]      text typed into cells into entry/[<tag>/]<name>.json
 *   npx tsx bench/windows/run.ts fit <name> [L]        columns fitted and texts typed into them, into fit/[<tag>/]<name>.json
 *   npx tsx bench/windows/run.ts cond <name>           conditional formats (in the set's locale) into cond/[<tag>/]<name>.json and cond/[<tag>/]<name>/
 *   npx tsx bench/windows/run.ts valid <name> [L]      data validation and outlines into valid/[<tag>/]<name>.json and valid/[<tag>/]<name>/
 *   npx tsx bench/windows/run.ts ui <name> [L]         a paste from the clipboard, and array cells through the object model, into ui/[<tag>/]<name>.json
 *   npx tsx bench/windows/run.ts ribbon <name> [L]     the ribbon's Number Format box over a ui set's books, by UI Automation in a session on a screen, into ribbon/[<tag>/]<name>.json
 *   npx tsx bench/windows/run.ts functions L           the formula cases the Mac has answers for in that locale
 *   npx tsx bench/windows/run.ts books L               the books the Mac has answers for in that locale (display, spill, ...)
 *   npx tsx bench/windows/run.ts all L                 both
 *
 * A locale L is `en`, `ru` or a tag of the package's locales (`bench/coverage/locales.ts`), en-US when left out; en-US's
 * answers are at the top of their directory, another locale's in a directory of its tag. Each command holds the
 * machine's lock and sets its locale first (restarting it when the system locale changes). FCOV_FRESH=1 runs every
 * case again; XLWIN_HOST names another machine.
 */
import { existsSync, readdirSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { stepDisplay, stepSpill, stepStruct, stepTables, stepWorkbook } from './books.js'
import { localeDir, parseTag } from '../coverage/locales.js'
import { stepFunctions } from './functions.js'
import { stepXlsxFilters, stepXlsxFormats } from './xlsx.js'
import { stepParts } from './parts.js'
import { stepHidden } from './hidden.js'
import { stepLinks } from './links.js'
import { stepLines } from './lines.js'
import { stepEntry } from './entry.js'
import { stepFit } from './fit.js'
import { stepCond } from './cond.js'
import type { CondSet } from '../coverage/cond/cases.js'
import { stepValid } from './valid.js'
import { stepUi } from './ui.js'
import { stepRibbon } from './ribbon.js'
import { env, withWindowsExcel, type Tag, type WinEnv } from './remote.js'

const ROOT = join(fileURLToPath(new URL('.', import.meta.url)), '../..')
/** The sets with Mac answers, by their files in `test/excel/<dir>`; none where the directory is not there. */
const macSets = (dir: string) => (existsSync(join(ROOT, 'test/excel', dir)) ? readdirSync(join(ROOT, 'test/excel', dir)).filter((f) => f.endsWith('.json')).map((f) => f.slice(0, -5)) : [])

const BOOK_STEPS = { spill: stepSpill, workbook: stepWorkbook, struct: stepStruct, parts: stepParts, hidden: stepHidden, lines: stepLines, links: stepLinks } as const
/** The parts sets: every set of `bench/coverage/parts/`, `threads` among them, which Windows alone runs. */
const PARTS_SETS = readdirSync(join(ROOT, 'bench/coverage/parts')).filter((f) => f.endsWith('.ts') && f !== 'cases.ts' && f !== 'predict.ts').map((f) => f.slice(0, -3))

/** The formula cases with Mac answers in a locale: the measurement, the case sets, in en-US the probes and the legacy sets. */
async function functions(tag: Tag, e: WinEnv) {
  const en = tag === 'en-US'
  const legacy = macSets('legacy').filter((s) => s !== 'measurement').map((s) => `legacy/${s}`)
  await stepFunctions(['measurement', ...(en ? ['probes', ...legacy] : []), ...macSets(`cases/${localeDir(tag)}`)], tag, e)
}

/** The books with Mac answers in a locale: what cells show, tables, data validation, and in en-US the spill, workbook, struct, parts, hidden and lines sets and the xlsx filters. */
async function books(tag: Tag, e: WinEnv) {
  const en = tag === 'en-US'
  for (const s of macSets(`display/${localeDir(tag)}`)) await stepDisplay(s, tag, e)
  for (const s of macSets(`tables/${localeDir(tag)}`)) await stepTables(s, tag, e)
  for (const s of macSets(`valid/${localeDir(tag)}`)) await stepValid(s, tag, e)
  if (!en) return
  // the order runs file keeps every order Excel gave: it is no set
  for (const s of macSets('spill').filter((s) => s !== 'order-runs')) await stepSpill(s, tag, e)
  for (const s of macSets('workbook')) await stepWorkbook(s, tag, e)
  for (const s of macSets('struct')) await stepStruct(s, tag, e)
  for (const s of PARTS_SETS) await stepParts(s, tag, e)
  for (const s of macSets('hidden')) await stepHidden(s, tag, e)
  for (const s of macSets('lines')) await stepLines(s, tag, e)
  await stepXlsxFilters(tag, e)
}

const [step, a, b] = process.argv.slice(2)
const t0 = Date.now()
if (step === 'env') console.log(env())
else if (step === 'measurement') await withWindowsExcel(parseTag(a), (e) => stepFunctions(['measurement'], parseTag(a), e))
else if (step === 'probes') await withWindowsExcel('en-US', (e) => stepFunctions(['probes'], 'en-US', e))
else if (step === 'set' && a) await withWindowsExcel(parseTag(b), (e) => stepFunctions([a], parseTag(b), e))
else if (step === 'legacy' && a) await withWindowsExcel('en-US', (e) => stepFunctions([`legacy/${a}`], 'en-US', e))
else if (step === 'display' && a) await withWindowsExcel(parseTag(b), (e) => stepDisplay(a, parseTag(b), e))
else if (step === 'tables' && a) await withWindowsExcel(parseTag(b), (e) => stepTables(a, parseTag(b), e))
else if (step in BOOK_STEPS && a) await withWindowsExcel('en-US', (e) => BOOK_STEPS[step as keyof typeof BOOK_STEPS](a, 'en-US', e))
else if (step === 'xlsx' && a === 'filters') await withWindowsExcel('en-US', (e) => stepXlsxFilters('en-US', e))
else if (step === 'xlsx' && a === 'formats') await withWindowsExcel(parseTag(b), (e) => stepXlsxFormats(parseTag(b), e))
else if (step === 'entry' && a) await withWindowsExcel(parseTag(b), (e) => stepEntry(a, parseTag(b), e))
else if (step === 'fit' && a) await withWindowsExcel(parseTag(b), (e) => stepFit(a, parseTag(b), e))
else if (step === 'cond' && a) {
  const tag = ((await import(`../coverage/cond/${a}.js`)) as CondSet).buildCondCases().locale ?? 'en-US'
  await withWindowsExcel(tag, (e) => stepCond(a, tag, e))
}
else if (step === 'valid' && a) await withWindowsExcel(parseTag(b), (e) => stepValid(a, parseTag(b), e))
else if (step === 'ui' && a) await withWindowsExcel(parseTag(b), (e) => stepUi(a, parseTag(b), e))
else if (step === 'ribbon' && a) await withWindowsExcel(parseTag(b), (e) => stepRibbon(a, parseTag(b), e))
else if (step === 'functions') await withWindowsExcel(parseTag(a), (e) => functions(parseTag(a), e))
else if (step === 'books') await withWindowsExcel(parseTag(a), (e) => books(parseTag(a), e))
else if (step === 'all') await withWindowsExcel(parseTag(a), async (e) => (await functions(parseTag(a), e), await books(parseTag(a), e)))
else {
  console.error('step: env | measurement <locale> | probes | set <name> [locale] | legacy <name> | display <name> [locale] | tables <name> [locale] | spill|workbook|struct|parts|hidden|lines|links <name> | xlsx filters | xlsx formats [locale] | entry <name> [locale] | fit <name> [locale] | cond <name> | valid <name> [locale] | ui <name> [locale] | ribbon <name> [locale] | functions|books|all <locale>; a locale is en, ru or a tag')
  process.exit(2)
}
console.log(`done in ${((Date.now() - t0) / 1000).toFixed(1)} s`)
