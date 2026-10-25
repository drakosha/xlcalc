/**
 * Acceptance against Excel for Windows (`test/excel/win/`, written by `bench/windows/`): the coverage measurement, the
 * probes and the case sets, each computed by the engine in the book the harness sent (en-US, or with the `locale` of
 * the answers, by the same tags and files as the Mac's: `bench/coverage/locales.ts`) and compared with Windows's answers
 * as with the Mac's. Where the engine knowingly differs from Excel for Windows,
 * `WIN_KNOWN` pins its answer with a reason; most of those are places where Excel for Windows and for Mac part.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import type { Val } from '../bench/coverage/formulajs.js'
import { localeDir, measurementFile, TAGS } from '../bench/coverage/locales.js'
import { buildProbes, PROBE_SHEETS, type Probe, type SheetCells } from '../bench/coverage/probes.js'
import type { BookInput } from '../src/contract.js'
import { compute, FUNCTION_NAMES } from '../src/index.js'
import { samePin } from '../bench/coverage/compare.js'
import { CASES, CLOCK, excelBook, matches, REASONS, show, toVal } from './excel-cases.js'
import { LOCALE_REASONS } from './excel-locale-cases.js'
import { caseBook, loadSet } from './excel-sets.js'
import { WIN_KNOWN_CASES } from './excel-win-known.js'
import './locales.js'

/** A stored Windows run of formula cases; `rejected` lists the cases whose formula Excel for Windows did not take. */
export interface WinSet {
  /** `measurement`, `probes` or a set's name. */
  name: string
  /** The locale the harness computed it in. */
  tag: string
  meta: Record<string, string | number | boolean>
  values: Record<string, Val>
  rejected?: string[]
  /** The cases as the harness sent them (the measurement's are `CASES`). */
  cases?: Probe[]
}

const dir = new URL('./excel/win/', import.meta.url)
const read = (rel: string) => JSON.parse(readFileSync(new URL(rel, dir), 'utf8')) as Omit<WinSet, 'name' | 'tag'>
const jsons = (rel: string) => (existsSync(new URL(rel, dir)) ? readdirSync(new URL(rel, dir)).filter((f) => f.endsWith('.json')).sort() : [])

/** Every stored Windows run of formula cases: the measurement and the sets by locale, the probes in en-US. */
export const WIN_SETS: WinSet[] = [
  ...TAGS.flatMap((tag) => (existsSync(new URL(`coverage/${measurementFile(tag)}`, dir)) ? [{ name: 'measurement', tag, ...read(`coverage/${measurementFile(tag)}`) }] : [])),
  ...(existsSync(new URL('probes.json', dir)) ? [{ name: 'probes', tag: 'en-US', ...read('probes.json') }] : []),
  ...TAGS.flatMap((tag) => jsons(`cases/${localeDir(tag)}`).map((f) => ({ name: f.slice(0, -5), tag, ...read(`cases/${localeDir(tag)}${f}`) }))),
]

/** Where Excel for Windows parts from Excel for Mac and the engine follows the Mac: the reasons of `WIN_KNOWN` that the Mac's lists have not. */
export const WIN_REASONS = {
  winRegexMatch: 'Match mode 3 of `XMATCH` and `XLOOKUP`, regular expressions: Excel for Windows 16.0.20430 has none and answers `#VALUE!`; Excel for Mac 16.114 and the engine find the pattern (`XMATCH("[A-Z][a-z]+y",A2:A9,3)`, IronCalc\'s `XMATCH.xlsx`).',
  winLongDate: '`[$-F800]` in ru-RU is the long date of the system: on Windows `31 января 2025 г.`, without the weekday; on the Mac and here `пятница, 31 января 2025 г.`.',
  winWeeknum: '`WEEKNUM` with return type 22, which Microsoft does not list: Excel for Windows answers a week (`WEEKNUM(45658,22)` is 1), Excel for Mac and the engine `#NUM!`.',
  winWeibull: '`WEIBULL` and `WEIBULL.DIST` with a scale of 1E-300: `#NUM!` in Excel for Windows, `#DIV/0!` in Excel for Mac and here.',
  winNameValue:
    'A formula over a name the book does not define is `#NAME?` as a whole, as an invalid formula; Excel gives the name\'s `#NAME?` to the function over it as a value, so `TYPE(falſe)`, a name in Excel for Windows, is 16 there.',
  winMath:
    'Trigonometry where the last bits decide: Excel for Windows computes `TAN(PI()/2)` as 16324552277619072, Excel for Mac and the engine 1.63312393531954E+16; the same for `SEC`, `CSC`, `COT` at their poles and for `IMTAN`, `IMCOS`, `IMSIN`, `IMPOWER` of large arguments, which Windows gives with other digits from the 13th on.',
} as const

export type Reason = keyof typeof WIN_REASONS | keyof typeof REASONS | keyof typeof LOCALE_REASONS

export const winReasonText = (r: Reason): string =>
  r in WIN_REASONS ? WIN_REASONS[r as keyof typeof WIN_REASONS] : r in LOCALE_REASONS ? LOCALE_REASONS[r as keyof typeof LOCALE_REASONS] : REASONS[r as keyof typeof REASONS]

/**
 * Cases where the engine knowingly differs from Excel for Windows, by `<locale>:<source>:<case id>` (`measurement`,
 * `probes` or a set's name): the engine's answer pinned instead of Windows's. An entry whose case matches fails.
 */
export const WIN_KNOWN: Readonly<Record<string, { reason: Reason; engine: Val }>> = WIN_KNOWN_CASES

/** A case passes: it matches Excel for Windows, or it is in WIN_KNOWN, differs and gives the pinned answer (`samePin`). */
export function winVerdict(key: string, fn: string | undefined, win: Val, engine: Val): { ok: boolean; why?: string } {
  const known = WIN_KNOWN[key]
  const same = matches(fn, win, engine)
  if (!known) return { ok: same }
  if (same) return { ok: false, why: 'matches Excel for Windows, remove it from WIN_KNOWN' }
  return samePin(known.engine, engine) ? { ok: true } : { ok: false, why: `WIN_KNOWN pins ${show(known.engine)}` }
}

export interface WinOutcome {
  key: string
  id: string
  fn?: string
  group: string
  formula: string
  win: Val
  engine: Val
  ok: boolean
  why?: string
}

const inTag = (book: BookInput, tag: string): BookInput => (tag === 'en-US' ? book : { ...book, locale: tag })

/** The cases a stored run has, as the harness sent them, and the book that holds them. */
async function casesOf(set: WinSet): Promise<{ cases: Probe[]; book: BookInput; current: boolean }> {
  if (set.name === 'measurement') {
    // as the Mac's measurement tests: the whitelisted functions' cases, in one book
    const whitelisted = CASES.filter((c) => FUNCTION_NAMES.includes(c.fn))
    return { cases: whitelisted.map((c) => ({ id: c.id, group: c.fn, fn: c.fn, formula: c.formula })), book: excelBook(whitelisted), current: true }
  }
  const mod: { buildProbes: () => Probe[]; SHEETS?: SheetCells } = set.name === 'probes' ? { buildProbes, SHEETS: PROBE_SHEETS } : await loadSet(set.name)
  const cases = mod.buildProbes()
  return { cases, book: caseBook(cases, mod.SHEETS), current: JSON.stringify(set.cases) === JSON.stringify(cases) }
}

/** A stored run computed by the engine in one book of its locale, and compared case by case (cases Windows did not take are left out). */
export async function runWinSet(set: WinSet): Promise<{ current: boolean; missing: string[]; outcomes: WinOutcome[] }> {
  const { cases, book, current } = await casesOf(set)
  const cells = compute(inTag(book, set.tag), { clock: CLOCK }).sheets.at(-1)!.cells
  const rejected = new Set(set.rejected ?? [])
  const outcomes = cases.flatMap((p, i): WinOutcome[] => {
    if (rejected.has(p.id) || !set.values[p.id]) return []
    const key = `${set.tag}:${set.name}:${p.id}`
    const engine = toVal(cells[`A${i + 1}`].value)
    const v = winVerdict(key, p.fn, set.values[p.id], engine)
    return [{ key, id: p.id, fn: p.fn, group: p.fn ?? p.group, formula: p.formula, win: set.values[p.id], engine, ok: v.ok, why: v.why }]
  })
  return { current, missing: cases.filter((p) => !set.values[p.id] && !rejected.has(p.id)).map((p) => p.id), outcomes }
}

export const describeWinOff = (o: WinOutcome) => `${o.id} ${o.formula}: Windows ${show(o.win)}, engine ${show(o.engine)}${o.why ? ` (${o.why})` : ''}`
