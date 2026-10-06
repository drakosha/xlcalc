/**
 * Acceptance against Excel ru-RU (Excel 16.113 for Mac with the Russian interface and locale): the coverage
 * measurement's cases computed in a book with `locale: 'ru-RU'` against `test/excel/coverage/excel-ru.json`,
 * and the case sets computed in ru-RU (`test/excel/cases/ru-RU/`, `run.ts probes <set> ru`). The clock is in
 * 2026, the year Excel computed them in, so date text without a year is in 2026. Comparison as for en-US;
 * known differences are pinned here with their reasons, as `KNOWN` does for en-US.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import type { Val } from '../bench/coverage/formulajs.js'
import type { BookInput } from '../src/contract.js'
import { compute } from '../src/index.js'
import { CASES, CLOCK, excelBook, matches, REASONS, show, toVal, type ExcelCase } from './excel-cases.js'
import { caseBook, loadSet, type StoredSet } from './excel-sets.js'

export const RU_EXCEL: Record<string, Val> = (JSON.parse(readFileSync(new URL('./excel/coverage/excel-ru.json', import.meta.url), 'utf8')) as { values: Record<string, Val> }).values

const dir = new URL('./excel/cases/ru-RU/', import.meta.url)

/** The case sets computed in Excel ru-RU. */
export const RU_SETS: StoredSet[] = existsSync(dir)
  ? readdirSync(dir)
      .filter((f) => f.endsWith('.json'))
      .sort()
      .map((f) => ({ name: f.slice(0, -5), ...(JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as Omit<StoredSet, 'name'>) }))
  : []

/** Reasons of differences that only ru-RU cases show; each is a line of the README's "Known differences from Excel". */
export const RU_REASONS = {
  nullError: '`#NULL!` is read as `#ERROR!`.',
  textFormats:
    '`TEXT` reads the format codes of en-US and of ru-RU; exponent formats (`0.00E+00`, `0,00E+00`), colours, elapsed time and bracketed conditions (`[h]:mm`, `[ч]:мм`) give `#VALUE!`.',
  secondDecimal:
    'A second decimal separator among the decimals of a number format (`#,##0,00` in ru-RU) is a literal between digits; Excel shows one digit fewer before it in some numbers (`TEXT(-0.5,"#,##0,00")` is `-,50,00` in Excel and `-,500,00` here).',
} as const

type Reason = keyof typeof REASONS | keyof typeof RU_REASONS

/** Cases where this engine knowingly differs from Excel ru-RU, by `<source>:<case id>`: the engine's answer pinned instead. */
export const RU_KNOWN: Readonly<Record<string, { reason: Reason; engine: Val }>> = {
  'measurement:CHIINV#7': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 10000232617.1965 } },
  'measurement:CHISQ.INV#14': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 10000208709.154 } },
  'measurement:CHISQ.INV.RT#15': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 10000232617.1965 } },
  'measurement:GAMMAINV#8': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 19999701973.2484 } },
  'locale:text#40': { reason: 'nullError', engine: { t: 's', v: '#ДЕЛ/0!; #ЗНАЧ!; #ССЫЛКА!; #ИМЯ?; #ЧИСЛО!; #ERROR!' } },
  'locale:format#24': { reason: 'textFormats', engine: { t: 'e', v: 'VALUE' } },
  'locale:format#93': { reason: 'textFormats', engine: { t: 'e', v: 'VALUE' } },
  'locale:format#94': { reason: 'textFormats', engine: { t: 'e', v: 'VALUE' } },
  'locale:format#95': { reason: 'textFormats', engine: { t: 'e', v: 'VALUE' } },
  'locale:format#96': { reason: 'textFormats', engine: { t: 'e', v: 'VALUE' } },
  'locale:format#244': { reason: 'secondDecimal', engine: { t: 's', v: '-,500,00' } },
}

/** The text of a reason, en-US's or ru-RU's own. */
export const reasonText = (r: Reason): string => (r in RU_REASONS ? RU_REASONS[r as keyof typeof RU_REASONS] : REASONS[r as keyof typeof REASONS])

/** A case passes: it matches Excel ru-RU, or it is in RU_KNOWN, differs from Excel and gives exactly the pinned answer. */
export function ruVerdict(key: string, fn: string | undefined, excel: Val, engine: Val): { ok: boolean; why?: string } {
  const known = RU_KNOWN[key]
  const same = matches(fn, excel, engine)
  if (!known) return { ok: same }
  if (same) return { ok: false, why: 'matches Excel, remove it from RU_KNOWN' }
  return engine.t === known.engine.t && engine.v === known.engine.v ? { ok: true } : { ok: false, why: `RU_KNOWN pins ${show(known.engine)}` }
}

const inRu = (book: BookInput): BookInput => ({ ...book, locale: 'ru-RU' })

export interface RuOutcome {
  id: string
  fn?: string
  formula: string
  excel: Val
  engine: Val
  ok: boolean
  why?: string
}

/** The measurement's cases computed in one ru-RU book. */
export function runRuCases(cases: readonly ExcelCase[] = CASES): RuOutcome[] {
  const cells = compute(inRu(excelBook(cases)), { clock: CLOCK }).sheets.at(-1)!.cells
  return cases.map((c, i) => {
    const engine = toVal(cells[`A${i + 1}`].value)
    const v = ruVerdict(`measurement:${c.id}`, c.fn, RU_EXCEL[c.id], engine)
    return { id: c.id, fn: c.fn, formula: c.formula, excel: RU_EXCEL[c.id], engine, ok: v.ok, why: v.why }
  })
}

/** A ru-RU case set computed in one ru-RU book, with its generator's current cases. */
export async function runRuSet(set: StoredSet): Promise<{ current: boolean; outcomes: (RuOutcome & { group: string })[] }> {
  const mod = await loadSet(set.name)
  const probes = mod.buildProbes()
  const cells = compute(inRu(caseBook(probes, mod.SHEETS)), { clock: CLOCK }).sheets.at(-1)!.cells
  const outcomes = probes.map((p, i) => {
    const engine = toVal(cells[`A${i + 1}`].value)
    const v = ruVerdict(`${set.name}:${p.id}`, p.fn, set.values[p.id], engine)
    return { id: p.id, fn: p.fn, group: p.fn ?? p.group, formula: p.formula, excel: set.values[p.id], engine, ok: v.ok, why: v.why }
  })
  return { current: JSON.stringify(set.cases) === JSON.stringify(probes) && probes.every((p) => set.values[p.id]), outcomes }
}

export const describeOff = (o: RuOutcome) => `${o.id} ${o.formula}: Excel ${show(o.excel)}, engine ${show(o.engine)}${o.why ? ` (${o.why})` : ''}`
