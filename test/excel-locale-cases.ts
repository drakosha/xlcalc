/**
 * Acceptance against Excel in each of the package's locales but en-US (en-US's is `excel-cases.ts` and `excel-sets.ts`),
 * by tag: the coverage measurement's cases computed in a book of the locale against Excel's answers in it
 * (`test/excel/coverage/excel-ru.json` for ru-RU, `measurementFile`), and the case sets computed in it
 * (`test/excel/cases/<tag>/`, `run.ts probes <set> <locale>`). Excel 16.113 for Mac computed them with the locale's
 * interface and region; a locale Excel has not computed in has no answers here, and the tests name it as skipped. The
 * clock is in 2026, the year Excel computed them in, so date text without a year is in 2026. Comparison as for en-US;
 * each locale's known differences are pinned in `KNOWN_BY_LOCALE`, as `KNOWN` does for en-US.
 */
import { existsSync, readFileSync } from 'node:fs'
import { samePin } from '../bench/coverage/compare.js'
import type { Val } from '../bench/coverage/formulajs.js'
import { localeDir, measurementFile, TAGS } from '../bench/coverage/locales.js'
import type { Probe, SheetCells } from '../bench/coverage/probes.js'
import type { BookInput } from '../src/contract.js'
import { compute } from '../src/index.js'
import { CASES, CLOCK, excelBook, matches, REASONS, show, toVal, type ExcelCase } from './excel-cases.js'
import { RU_KNOWN, RU_REASONS } from './excel-ru-known.js'
import { caseBook, loadSet, type StoredSet } from './excel-sets.js'
import { storedSets } from './stored-sets.js'
import './locales.js'

/** The locales compared here: the package's, but en-US. */
export const LOCALE_TAGS: readonly string[] = TAGS.filter((t) => t !== 'en-US')

/** Reasons of differences that only some locale's cases show, by name; each is a line of docs/limitations.md's "Known differences from Excel". */
export const LOCALE_REASONS = { ...RU_REASONS } as const

export type LocaleReason = keyof typeof REASONS | keyof typeof LOCALE_REASONS

/** A locale's known differences by `<source>:<case id>` (`measurement` or a case set's name): the engine's answer pinned instead of Excel's. */
export type KnownList = Readonly<Record<string, { reason: LocaleReason; engine: Val }>>

/** The known differences by tag; a locale without a list has none. An entry whose case matches Excel fails the tests. */
export const KNOWN_BY_LOCALE: Readonly<Record<string, KnownList>> = { 'ru-RU': RU_KNOWN }

/** The text of a reason, en-US's or a locale's own. */
export const reasonText = (r: LocaleReason): string => (r in LOCALE_REASONS ? LOCALE_REASONS[r as keyof typeof LOCALE_REASONS] : REASONS[r as keyof typeof REASONS])

/** Where the measurement's answers of a locale are stored, as messages name the file. */
export const measurementPath = (tag: string) => `test/excel/coverage/${measurementFile(tag)}`

const json = <T>(url: URL) => JSON.parse(readFileSync(url, 'utf8')) as T

/** Excel's answers to the measurement by locale, for the locales Excel computed it in. */
export const LOCALE_EXCEL: ReadonlyMap<string, Record<string, Val>> = new Map(
  LOCALE_TAGS.flatMap((tag) => {
    const file = new URL(`../${measurementPath(tag)}`, import.meta.url)
    return existsSync(file) ? [[tag, json<{ values: Record<string, Val> }>(file).values] as const] : []
  }),
)

/** The case sets computed in Excel in each locale (`test/excel/cases/<tag>/`), none where it has no directory. */
export const LOCALE_SETS: ReadonlyMap<string, StoredSet[]> = new Map(
  LOCALE_TAGS.map((tag) => [tag, storedSets<Omit<StoredSet, 'name'>>(`cases/${localeDir(tag)}`).map(({ dir, ...set }) => set)]),
)

/** A case passes: it matches Excel in the locale, or it is a known difference there, differs from Excel and gives the pinned answer (`samePin`). */
export function localeVerdict(tag: string, key: string, fn: string | undefined, excel: Val, engine: Val): { ok: boolean; why?: string } {
  const known = KNOWN_BY_LOCALE[tag]?.[key]
  const same = matches(fn, excel, engine)
  if (!known) return { ok: same }
  if (same) return { ok: false, why: `matches Excel, remove it from the known differences of ${tag}` }
  return samePin(known.engine, engine) ? { ok: true } : { ok: false, why: `the known differences of ${tag} pin ${show(known.engine)}` }
}

const inLocale = (book: BookInput, tag: string): BookInput => ({ ...book, locale: tag })

export interface LocaleOutcome {
  id: string
  fn?: string
  formula: string
  excel: Val
  engine: Val
  ok: boolean
  why?: string
}

/** The measurement's cases computed in one book of the locale, against Excel's answers there. */
export function runLocaleCases(tag: string, cases: readonly ExcelCase[] = CASES): LocaleOutcome[] {
  const excel = LOCALE_EXCEL.get(tag)
  if (!excel) throw new Error(`no answers of Excel ${tag} to the measurement (${measurementPath(tag)})`)
  const cells = compute(inLocale(excelBook(cases), tag), { clock: CLOCK }).sheets.at(-1)!.cells
  return cases.map((c, i) => {
    const engine = toVal(cells[`A${i + 1}`].value)
    const v = localeVerdict(tag, `measurement:${c.id}`, c.fn, excel[c.id], engine)
    return { id: c.id, fn: c.fn, formula: c.formula, excel: excel[c.id], engine, ok: v.ok, why: v.why }
  })
}

/** The engine's answers to a case set's cases, computed in one book of the locale as Excel computes them there. */
export function localeAnswers(tag: string, probes: readonly Probe[], sheets?: SheetCells): Val[] {
  const cells = compute(inLocale(caseBook(probes, sheets), tag), { clock: CLOCK }).sheets.at(-1)!.cells
  return probes.map((_, i) => toVal(cells[`A${i + 1}`].value))
}

/** A locale's case set computed in one book of the locale, with its generator's current cases. */
export async function runLocaleSet(tag: string, set: StoredSet): Promise<{ current: boolean; outcomes: (LocaleOutcome & { group: string })[] }> {
  const mod = await loadSet(set.name)
  const probes = mod.buildProbes()
  const answers = localeAnswers(tag, probes, mod.SHEETS)
  const outcomes = probes.map((p, i) => {
    const engine = answers[i]
    const v = localeVerdict(tag, `${set.name}:${p.id}`, p.fn, set.values[p.id], engine)
    return { id: p.id, fn: p.fn, group: p.fn ?? p.group, formula: p.formula, excel: set.values[p.id], engine, ok: v.ok, why: v.why }
  })
  return { current: JSON.stringify(set.cases) === JSON.stringify(probes) && probes.every((p) => set.values[p.id]), outcomes }
}

export const describeOff = (o: LocaleOutcome) => `${o.id} ${o.formula}: Excel ${show(o.excel)}, engine ${show(o.engine)}${o.why ? ` (${o.why})` : ''}`
