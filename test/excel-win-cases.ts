/**
 * Acceptance against Excel for Windows (`test/excel/win/`, written by `bench/windows/`): the coverage measurement, the
 * probes and the case sets, each computed by the engine in the book the harness sent (en-US, or `locale: 'ru-RU'`) and
 * compared with Windows's answers as with the Mac's. Where the engine knowingly differs from Excel for Windows,
 * `WIN_KNOWN` pins its answer with a reason; most of those are places where Excel for Windows and for Mac part.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import type { Val } from '../bench/coverage/formulajs.js'
import { buildProbes, PROBE_SHEETS, type Probe, type SheetCells } from '../bench/coverage/probes.js'
import type { BookInput } from '../src/contract.js'
import { compute } from '../src/index.js'
import { CASES, CLOCK, excelBook, matches, REASONS, show, toVal } from './excel-cases.js'
import { RU_REASONS } from './excel-ru-cases.js'
import { caseBook, loadSet } from './excel-sets.js'

export type Tag = 'en-US' | 'ru-RU'

/** A stored Windows run of formula cases; `rejected` lists the cases whose formula Excel for Windows did not take. */
export interface WinSet {
  /** `measurement`, `probes` or a set's name. */
  name: string
  tag: Tag
  meta: Record<string, string | number | boolean>
  values: Record<string, Val>
  rejected?: string[]
  /** The cases as the harness sent them (the measurement's are `CASES`). */
  cases?: Probe[]
}

const dir = new URL('./excel/win/', import.meta.url)
const read = (rel: string) => JSON.parse(readFileSync(new URL(rel, dir), 'utf8')) as Omit<WinSet, 'name' | 'tag'>
const jsons = (rel: string) => (existsSync(new URL(rel, dir)) ? readdirSync(new URL(rel, dir)).filter((f) => f.endsWith('.json')).sort() : [])

/** The stored Windows answer files of a kind (`display/`, `display/ru-RU/`, `spill/`, ...), each with its set's name. */
export const winFiles = <T>(sub: string): (T & { name: string })[] => jsons(sub).map((f) => ({ name: f.slice(0, -5), ...(JSON.parse(readFileSync(new URL(sub + f, dir), 'utf8')) as T) }))

/** Every stored Windows run of formula cases: the measurement and the sets by locale, the probes in en-US. */
export const WIN_SETS: WinSet[] = [
  ...(['en', 'ru'] as const).flatMap((l) => (existsSync(new URL(`coverage/excel-${l}.json`, dir)) ? [{ name: 'measurement', tag: (l === 'en' ? 'en-US' : 'ru-RU') as Tag, ...read(`coverage/excel-${l}.json`) }] : [])),
  ...(existsSync(new URL('probes.json', dir)) ? [{ name: 'probes', tag: 'en-US' as Tag, ...read('probes.json') }] : []),
  ...jsons('cases/').map((f) => ({ name: f.slice(0, -5), tag: 'en-US' as Tag, ...read(`cases/${f}`) })),
  ...jsons('cases/ru-RU/').map((f) => ({ name: f.slice(0, -5), tag: 'ru-RU' as Tag, ...read(`cases/ru-RU/${f}`) })),
]

/** Reasons that only Excel for Windows shows; each is a line of the README's "Excel for Windows" known differences. */
export const WIN_REASONS = {} as const

type Reason = keyof typeof WIN_REASONS | keyof typeof REASONS | keyof typeof RU_REASONS

export const winReasonText = (r: Reason): string =>
  r in WIN_REASONS ? WIN_REASONS[r as keyof typeof WIN_REASONS] : r in RU_REASONS ? RU_REASONS[r as keyof typeof RU_REASONS] : REASONS[r as keyof typeof REASONS]

/**
 * Cases where the engine knowingly differs from Excel for Windows, by `<locale>:<source>:<case id>` (`measurement`,
 * `probes` or a set's name): the engine's answer pinned instead of Windows's. An entry whose case matches fails.
 */
export const WIN_KNOWN: Readonly<Record<string, { reason: Reason; engine: Val }>> = {}

/** A case passes: it matches Excel for Windows, or it is in WIN_KNOWN, differs and gives exactly the pinned answer. */
export function winVerdict(key: string, fn: string | undefined, win: Val, engine: Val): { ok: boolean; why?: string } {
  const known = WIN_KNOWN[key]
  const same = matches(fn, win, engine)
  if (!known) return { ok: same }
  if (same) return { ok: false, why: 'matches Excel for Windows, remove it from WIN_KNOWN' }
  return engine.t === known.engine.t && engine.v === known.engine.v ? { ok: true } : { ok: false, why: `WIN_KNOWN pins ${show(known.engine)}` }
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

const inTag = (book: BookInput, tag: Tag): BookInput => (tag === 'ru-RU' ? { ...book, locale: 'ru-RU' } : book)

/** The cases a stored run has, as the harness sent them, and the book that holds them. */
async function casesOf(set: WinSet): Promise<{ cases: Probe[]; book: BookInput; current: boolean }> {
  if (set.name === 'measurement') {
    const cases = CASES.map((c) => ({ id: c.id, group: c.fn, fn: c.fn, formula: c.formula }))
    return { cases, book: excelBook(CASES), current: true }
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
