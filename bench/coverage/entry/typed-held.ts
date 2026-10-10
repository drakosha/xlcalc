/**
 * Held out from the set `typed`, predicted by the engine before Excel ran it: the euro sign, numbers past the range in
 * scientific form, years of one digit (in VALUE too), the formats formulas take, codes that are a locale's built-in
 * formats, built-in formats of other kinds, spaces around dates and times, more groups, dates and fractions.
 */
import type { NumberFormat } from '../../../src/contract.js'
import { collector, type EntryCase } from './cases.js'

const GENERAL: Record<string, string[]> = {
  euro: ['€1,234.5', '€1 234,5', '€ 5', '-€5', '€-5', '5€', '€5%', '€5,5', '€5.5', '(€5)', '€'],
  range: ['1,5E+400', '1.5E+400', '-1E+309', '1E-330', '0E+400', '1E+308'],
  year: ['1/2/3', '1.2.3', '15.3.6', '3/15/6', '1/2/345', '2026-3-5', '5-3-2026', '=VALUE("1-2-3")', '=VALUE("1/2/3")', '=ЗНАЧЕН("1-2-3")', '=ЗНАЧЕН("1.2.3")'],
  ffx: ['=TODAY()+1', '=1+TODAY()', '=YEAR(TODAY())', '=TODAY()-C2', '=EDATE(C2,1)', '=EOMONTH(C2,0)', '=DATEVALUE("3/15/2026")', '=TIMEVALUE("18:30")', '=NOW()-TODAY()', '=PV(0.1,12,100)', '=FV(0.1,12,100)', '=IPMT(0.1,1,12,1000)', '=NPV(0.1,100,200)', '=SUM(C2,1)', '=MAX(C2,C1)', '=C1+C2', '=IF(TRUE,TODAY(),0)', '=WORKDAY(C2,1)', '=C2*1', '=5%*2', '=DOLLAR(5)', '=-TODAY()', '=(TODAY())', '=SUM(TODAY())', '=TIME(1,0,0)+TODAY()', '=HOUR(NOW())', '=СЕГОДНЯ()+1', '=ДАТАМЕС(C2;1)', '=ПС(0,1;12;100)', '=ГОД(СЕГОДНЯ())', '=ЕСЛИ(ИСТИНА;СЕГОДНЯ();0)', '=МАКС(C2;C1)'],
  space: [' 15.03.2026', '15.03.2026 ', ' 3/15/2026', ' 18:30', '18:30 ', ' 5%', '  1 234,5', ' 1,234.5', ' #N/A', ' #Н/Д', '#знач!'],
  more: ['12,345,678', '1 234 567', '1,234.', '1 234,', '+$5', '+5 ₽', '2/3', '0 1/100', '1 1/3', '5 1/2%', '15 марта', '15 March 2026', 'сен 2026', '15 сент 2026', '1.1.2026 0:00', '9999-12-31', '1899-12-31', '1900-02-29'],
}

/** Codes that are built-in formats in one locale or both, and built-in formats of other kinds, by a short name for the ids. */
const FORMATS: Record<string, NumberFormat> = {
  'code-dmy': 'dd.mm.yyyy',
  'code-mdy': 'm/d/yyyy',
  'code-grp': '#,##0',
  'code-pct': '0%',
  b37: 37,
  b41: 41,
  b45: 45,
  b46: 46,
  b48: 48,
  b13: 13,
  b15: 15,
  b18: 18,
  b7: 7,
  b5: 5,
}

const INTO = ['5', '15.03.2026', '3/15/2026', '18:30', '5%', '1 1/2', '1,5E+3', '1.5E+3', '5 ₽', '$5', '=TODAY()', '=СЕГОДНЯ()']

export function buildEntryCases(): EntryCase[] {
  const { out, add } = collector()
  for (const [group, texts] of Object.entries(GENERAL)) for (const t of texts) add(group, t)
  for (const [name, format] of Object.entries(FORMATS)) for (const t of INTO) add(`in-${name}`, t, format)
  return out
}
