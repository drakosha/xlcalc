/**
 * The engine's prediction of the sets `value-names` and `value-names-reopen` in a language, written before Excel's
 * answers are read: each case played on the engine alone (`drive.ts`), from the rows or from the book `value-names`
 * saved, read; for each phase the cells of `PT1`'s area. Run: `npx tsx bench/coverage/pivot/value-names-predict.ts <set> [ru-RU]`; it writes
 * `bench/coverage/pivot/<set>[.ru-RU].predicted.json`. The comparison with Excel: `test/pivot-value-names.test.ts`.
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { registerLocale, registerPivot, type BookInput, type CellValue } from '../../../src/index.js'
import { ruRU } from '../../../src/locales/ru-RU.js'
import { pivotArea } from '../../../src/pivot/area.js'
import { pivotEngine } from '../../../src/pivot/index.js'
import { rectText } from '../../../src/structure/area.js'
import { readXlsx } from '../../../src/xlsx.js'
import type { PivotCase, PivotSet } from './cases.js'
import { drive } from './drive.js'
import { shown } from './phases.js'

const ROOT = new URL('../../../', import.meta.url)

/** The cells of `PT1`'s area on `P`. */
function area(book: BookInput): CellValue[][] {
  const t = book.sheets.find((s) => s.name === 'P')!.pivotTables!.find((x) => x.name === 'PT1')!
  return shown(book, 'P', rectText(pivotArea(t)!))
}

/** Each phase of a case, the cells of `PT1`'s area at its end; a case from a saved book starts from that book read. */
export async function predictedNames(c: PivotCase, locale?: string): Promise<CellValue[][][]> {
  registerPivot(pivotEngine)
  registerLocale(ruRU)
  // a saved book in the language Excel saved it in, which the file does not name
  const read = c.from ? (await readXlsx(readFileSync(new URL(`test/excel/win/pivot/${locale ? `${locale}/` : ''}${c.from}.xlsx`, ROOT)))).book : undefined
  const saved = read && locale ? { ...read, locale } : read
  return drive(c, saved, locale).map((p) => area(p.book))
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const [set, locale] = process.argv.slice(2)
  const { buildPivotCases } = (await import(`./${set}.ts`)) as PivotSet
  const out: Record<string, CellValue[][][]> = {}
  for (const c of buildPivotCases()) out[c.id] = await predictedNames(c, locale)
  writeFileSync(new URL(`${set}${locale ? `.${locale}` : ''}.predicted.json`, import.meta.url), JSON.stringify(out) + '\n')
  console.log(`${Object.keys(out).length} cases`)
}
