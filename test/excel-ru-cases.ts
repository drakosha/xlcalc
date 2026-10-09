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
  secondDecimal:
    'A second decimal separator among the decimals of a number format (`#,##0,00` in ru-RU) is a literal between digits; Excel shows one digit fewer before it in some numbers (`TEXT(-0.5,"#,##0,00")` is `-,50,00` in Excel and `-,500,00` here).',
  macCollation:
    'Excel for Mac puts Cyrillic before Latin in ru-RU (`"a"<"а"` is FALSE there); the engine puts Latin first, as Excel for Windows does in every locale, in comparisons, criteria, lookups and SORT.',
  macMonthShort:
    'Excel for Mac writes the short month names of ru-RU (`МММ`) `февр`, `март`, `июнь`, `июль`, `сент`, `нояб`; the engine takes three letters, as Excel for Windows does (`фев`, `мар`, `июн`, `июл`, `сен`, `ноя`).',
} as const

type Reason = keyof typeof REASONS | keyof typeof RU_REASONS

/** Cases where this engine knowingly differs from Excel ru-RU, by `<source>:<case id>`: the engine's answer pinned instead. */
export const RU_KNOWN: Readonly<Record<string, { reason: Reason; engine: Val }>> = {
  'measurement:CHAR#3': { reason: 'codePage', engine: { t: 's', v: 'И' } },
  'measurement:CODE#3': { reason: 'codePage', engine: { t: 'n', v: 198 } },
  'measurement:CODE#5': { reason: 'codePage', engine: { t: 'n', v: 207 } },
  'measurement:CODE#8': { reason: 'codePage', engine: { t: 'n', v: 200 } },
  'measurement:CODE#12': { reason: 'codePage', engine: { t: 'n', v: 207 } },
  'char-code:CODE#142': { reason: 'outsideCode', engine: { t: 'n', v: 63 } },
  'char-code:CODE#143': { reason: 'outsideCode', engine: { t: 'n', v: 63 } },
  'char-code:CODE#144': { reason: 'outsideCode', engine: { t: 'n', v: 63 } },
  'char-code:CODE#145': { reason: 'outsideCode', engine: { t: 'n', v: 63 } },
  'char-code:CODE#146': { reason: 'outsideCode', engine: { t: 'n', v: 63 } },
  'char-code:CODE#147': { reason: 'outsideCode', engine: { t: 'n', v: 63 } },
  'char-code:CODE#148': { reason: 'outsideCode', engine: { t: 'n', v: 63 } },
  'measurement:CHIINV#7': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 10000232617.1965 } },
  'measurement:CHISQ.INV#14': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 10000208709.154 } },
  'measurement:CHISQ.INV.RT#15': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 10000232617.1965 } },
  'measurement:GAMMAINV#8': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 19999701973.2484 } },
  'locale:format#244': { reason: 'secondDecimal', engine: { t: 's', v: '-,500,00' } },
  'measurement:SORT#9': { reason: 'macCollation', engine: { t: 's', v: "x*y" } },
  'measurement:SORT#10': { reason: 'macCollation', engine: { t: 's', v: "Арeнда*" } },
  'measurement:SORT#11': { reason: 'macCollation', engine: { t: 's', v: "Чай" } },
  'collation:compare#1': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#3': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#5': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#6': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#7': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#10': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#11': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#12': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#13': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#15': { reason: 'macCollation', engine: { t: 'b', v: false } },
  'collation:compare#17': { reason: 'macCollation', engine: { t: 'b', v: false } },
  'collation:compare#21': { reason: 'macCollation', engine: { t: 'b', v: false } },
  'collation:compare#22': { reason: 'macCollation', engine: { t: 'b', v: false } },
  'collation:compare#28': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#30': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#31': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#32': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#35': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#36': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#37': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#38': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#40': { reason: 'macCollation', engine: { t: 'b', v: false } },
  'collation:compare#44': { reason: 'macCollation', engine: { t: 'b', v: false } },
  'collation:compare#45': { reason: 'macCollation', engine: { t: 'b', v: false } },
  'collation:compare#51': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#52': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#53': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#56': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#57': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#58': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#59': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#63': { reason: 'macCollation', engine: { t: 'b', v: false } },
  'collation:compare#64': { reason: 'macCollation', engine: { t: 'b', v: false } },
  'collation:compare#71': { reason: 'macCollation', engine: { t: 'b', v: false } },
  'collation:compare#72': { reason: 'macCollation', engine: { t: 'b', v: false } },
  'collation:compare#78': { reason: 'macCollation', engine: { t: 'b', v: false } },
  'collation:compare#79': { reason: 'macCollation', engine: { t: 'b', v: false } },
  'collation:compare#86': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#87': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#88': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#89': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#91': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#92': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#93': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:compare#94': { reason: 'macCollation', engine: { t: 'b', v: true } },
  'collation:sort#4': { reason: 'macCollation', engine: { t: 's', v: "A" } },
  'collation:sort#5': { reason: 'macCollation', engine: { t: 's', v: "ab" } },
  'collation:sort#6': { reason: 'macCollation', engine: { t: 's', v: "abc" } },
  'collation:sort#7': { reason: 'macCollation', engine: { t: 's', v: "Abc" } },
  'collation:sort#8': { reason: 'macCollation', engine: { t: 's', v: "b" } },
  'collation:sort#9': { reason: 'macCollation', engine: { t: 's', v: "e" } },
  'collation:sort#10': { reason: 'macCollation', engine: { t: 's', v: "é" } },
  'collation:sort#11': { reason: 'macCollation', engine: { t: 's', v: "ea" } },
  'collation:sort#12': { reason: 'macCollation', engine: { t: 's', v: "éa" } },
  'collation:sort#13': { reason: 'macCollation', engine: { t: 's', v: "x1" } },
  'collation:sort#14': { reason: 'macCollation', engine: { t: 's', v: "z" } },
  'collation:sort#15': { reason: 'macCollation', engine: { t: 's', v: "Z" } },
  'collation:sort#16': { reason: 'macCollation', engine: { t: 's', v: "а" } },
  'collation:sort#17': { reason: 'macCollation', engine: { t: 's', v: "аб" } },
  'collation:sort#18': { reason: 'macCollation', engine: { t: 's', v: "абв" } },
  'collation:sort#19': { reason: 'macCollation', engine: { t: 's', v: "Абв" } },
  'collation:sort#20': { reason: 'macCollation', engine: { t: 's', v: "Б" } },
  'collation:sort#21': { reason: 'macCollation', engine: { t: 's', v: "е" } },
  'collation:sort#22': { reason: 'macCollation', engine: { t: 's', v: "ё" } },
  'collation:sort#23': { reason: 'macCollation', engine: { t: 's', v: "Ё" } },
  'collation:sort#24': { reason: 'macCollation', engine: { t: 's', v: "ж" } },
  'collation:sort#25': { reason: 'macCollation', engine: { t: 's', v: "и" } },
  'collation:sort#26': { reason: 'macCollation', engine: { t: 's', v: "й" } },
  'collation:sort#27': { reason: 'macCollation', engine: { t: 's', v: "х1" } },
  'collation:sort#28': { reason: 'macCollation', engine: { t: 's', v: "ъ" } },
  'collation:sort#29': { reason: 'macCollation', engine: { t: 's', v: "ь" } },
  'collation:sort#30': { reason: 'macCollation', engine: { t: 's', v: "я" } },
  'collation:sort#31': { reason: 'macCollation', engine: { t: 's', v: "я" } },
  'collation:sort#32': { reason: 'macCollation', engine: { t: 's', v: "х1" } },
  'collation:sort#33': { reason: 'macCollation', engine: { t: 's', v: "ж" } },
  'collation:sort#34': { reason: 'macCollation', engine: { t: 's', v: "е" } },
  'collation:sort#35': { reason: 'macCollation', engine: { t: 's', v: "Абв" } },
  'collation:sort#36': { reason: 'macCollation', engine: { t: 's', v: "z" } },
  'collation:sort#37': { reason: 'macCollation', engine: { t: 's', v: "éa" } },
  'collation:sort#38': { reason: 'macCollation', engine: { t: 's', v: "e" } },
  'collation:sort#39': { reason: 'macCollation', engine: { t: 's', v: "Abc" } },
  'collation:criteria#1': { reason: 'macCollation', engine: { t: 'n', v: 26 } },
  'collation:criteria#2': { reason: 'macCollation', engine: { t: 'n', v: 7 } },
  'collation:criteria#3': { reason: 'macCollation', engine: { t: 'n', v: 29 } },
  'collation:criteria#4': { reason: 'macCollation', engine: { t: 'n', v: 15 } },
  'collation:criteria#5': { reason: 'macCollation', engine: { t: 'n', v: 21 } },
  'collation:criteria#6': { reason: 'macCollation', engine: { t: 'n', v: 21 } },
  'collation:criteria#7': { reason: 'macCollation', engine: { t: 'n', v: 9 } },
  'collation:criteria#8': { reason: 'macCollation', engine: { t: 'n', v: 15 } },
  'collation:criteria#9': { reason: 'macCollation', engine: { t: 'n', v: 19 } },
  'locale:format#7': { reason: 'macMonthCase', engine: { t: 's', v: "Январь" } },
  'locale:format#8': { reason: 'macMonthCase', engine: { t: 's', v: "Я" } },
  'locale:format#11': { reason: 'macMonthCase', engine: { t: 's', v: "31 Январь 2025" } },
  'locale:format#12': { reason: 'macMonthCase', engine: { t: 's', v: "Январь 2025" } },
  'locale:format#267': { reason: 'macMonthCase', engine: { t: 's', v: "Январь" } },
  'locale:format#269': { reason: 'macMonthCase', engine: { t: 's', v: "15 Январь" } },
  'locale:format#270': { reason: 'macMonthCase', engine: { t: 's', v: "Я" } },
  'locale:format#271': { reason: 'macMonthCase', engine: { t: 's', v: "Февраль" } },
  'locale:format#272': { reason: 'macMonthShort', engine: { t: 's', v: "фев" } },
  'locale:format#273': { reason: 'macMonthCase', engine: { t: 's', v: "15 Февраль" } },
  'locale:format#274': { reason: 'macMonthCase', engine: { t: 's', v: "Ф" } },
  'locale:format#275': { reason: 'macMonthCase', engine: { t: 's', v: "Март" } },
  'locale:format#276': { reason: 'macMonthShort', engine: { t: 's', v: "мар" } },
  'locale:format#277': { reason: 'macMonthCase', engine: { t: 's', v: "15 Март" } },
  'locale:format#278': { reason: 'macMonthCase', engine: { t: 's', v: "М" } },
  'locale:format#279': { reason: 'macMonthCase', engine: { t: 's', v: "Апрель" } },
  'locale:format#281': { reason: 'macMonthCase', engine: { t: 's', v: "15 Апрель" } },
  'locale:format#282': { reason: 'macMonthCase', engine: { t: 's', v: "А" } },
  'locale:format#283': { reason: 'macMonthCase', engine: { t: 's', v: "Май" } },
  'locale:format#285': { reason: 'macMonthCase', engine: { t: 's', v: "15 Май" } },
  'locale:format#286': { reason: 'macMonthCase', engine: { t: 's', v: "М" } },
  'locale:format#287': { reason: 'macMonthCase', engine: { t: 's', v: "Июнь" } },
  'locale:format#288': { reason: 'macMonthShort', engine: { t: 's', v: "июн" } },
  'locale:format#289': { reason: 'macMonthCase', engine: { t: 's', v: "15 Июнь" } },
  'locale:format#290': { reason: 'macMonthCase', engine: { t: 's', v: "И" } },
  'locale:format#291': { reason: 'macMonthCase', engine: { t: 's', v: "Июль" } },
  'locale:format#292': { reason: 'macMonthShort', engine: { t: 's', v: "июл" } },
  'locale:format#293': { reason: 'macMonthCase', engine: { t: 's', v: "15 Июль" } },
  'locale:format#294': { reason: 'macMonthCase', engine: { t: 's', v: "И" } },
  'locale:format#295': { reason: 'macMonthCase', engine: { t: 's', v: "Август" } },
  'locale:format#297': { reason: 'macMonthCase', engine: { t: 's', v: "15 Август" } },
  'locale:format#298': { reason: 'macMonthCase', engine: { t: 's', v: "А" } },
  'locale:format#299': { reason: 'macMonthCase', engine: { t: 's', v: "Сентябрь" } },
  'locale:format#300': { reason: 'macMonthShort', engine: { t: 's', v: "сен" } },
  'locale:format#301': { reason: 'macMonthCase', engine: { t: 's', v: "15 Сентябрь" } },
  'locale:format#302': { reason: 'macMonthCase', engine: { t: 's', v: "С" } },
  'locale:format#303': { reason: 'macMonthCase', engine: { t: 's', v: "Октябрь" } },
  'locale:format#305': { reason: 'macMonthCase', engine: { t: 's', v: "15 Октябрь" } },
  'locale:format#306': { reason: 'macMonthCase', engine: { t: 's', v: "О" } },
  'locale:format#307': { reason: 'macMonthCase', engine: { t: 's', v: "Ноябрь" } },
  'locale:format#308': { reason: 'macMonthShort', engine: { t: 's', v: "ноя" } },
  'locale:format#309': { reason: 'macMonthCase', engine: { t: 's', v: "15 Ноябрь" } },
  'locale:format#310': { reason: 'macMonthCase', engine: { t: 's', v: "Н" } },
  'locale:format#311': { reason: 'macMonthCase', engine: { t: 's', v: "Декабрь" } },
  'locale:format#313': { reason: 'macMonthCase', engine: { t: 's', v: "15 Декабрь" } },
  'locale:format#314': { reason: 'macMonthCase', engine: { t: 's', v: "Д" } },
  'locale:criteria#30': { reason: 'macCollation', engine: { t: 'n', v: 6 } },
  'locale:criteria#31': { reason: 'macCollation', engine: { t: 'n', v: 3 } },
  'text-format-ru-more:TEXT#11': { reason: 'macMonthCase', engine: { t: 's', v: "31 Январь 2025" } },
  'text-format-ru-more:TEXT#12': { reason: 'macMonthCase', engine: { t: 's', v: "31 Январь" } },
  'text-format-ru:date#55': { reason: 'macMonthCase', engine: { t: 's', v: "Январь" } },
  'text-format-ru:date#56': { reason: 'macMonthCase', engine: { t: 's', v: "Январь" } },
  'text-format-ru:date#58': { reason: 'macMonthCase', engine: { t: 's', v: "Февраль" } },
  'text-format-ru:date#59': { reason: 'macMonthCase', engine: { t: 's', v: "Январь" } },
  'text-format-ru:date#60': { reason: 'macMonthCase', engine: { t: 's', v: "Январь" } },
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
