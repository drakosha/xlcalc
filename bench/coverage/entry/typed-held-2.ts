/**
 * Held out from the sets `typed` and `typed-held`, predicted by the engine after their rules were taken in and before
 * Excel ran it: formats of formulas, codes of built-in formats, the euro sign, percent with a fraction, the limits of
 * scientific form, spaces around dates.
 */
import type { NumberFormat } from '../../../src/contract.js'
import { collector, type EntryCase } from './cases.js'

const GENERAL: Record<string, string[]> = {
  ffx2: ['=TODAY()-1', '=TODAY()-1.5', '=SUM(1,TODAY())', '=MAX(1,TODAY())', '=MIN(TODAY(),1)', '=DATE(2026,1,1)+TIME(1,0,0)', '=TODAY()+TIME(1,0,0)', '=NOW()+1', '=-NOW()', '=(TIME(1,0,0))', '=PMT(0.1,12,1000)*2', '=-PMT(0.1,12,1000)', '=ROUND(TODAY(),0)', '=TODAY()*1', '=PPMT(0.1,1,12,1000)', '=TODAY()&""', '=AVERAGE(TODAY(),1)', '=ДАТА(2026;1;1)+ВРЕМЯ(1;0;0)', '=СУММ(1;СЕГОДНЯ())', '=ПЛТ(0,1;12;1000)*2'],
  euro2: ['€1,5', '€1.5', '€ -5', '-5€', '(5 €)', '€€5', '€5 €'],
  pct2: ['1/2%', '-1/4%', '2 1/2%', '1/0%', '%1/2'],
  sci2: ['1E+307', '9.99E+307', '9,99E+307', '-1E+308', '1E-309', '0E-500', '1E+3%'],
  space2: ['  15.03.2026', '15.03.2026  ', '  3/15/2026', '3/15/2026  ', ' 1/2', ' 6 PM', '6 PM '],
}

/** Codes of built-in formats in both locales or one, by a short name for the ids. */
const FORMATS: Record<string, NumberFormat> = {
  'code-sci': '0.00E+00',
  'code-frac': '# ?/?',
  'code-mmss': 'mm:ss',
  'code-pct2': '0.00%',
  'code-usd': '$#,##0_);($#,##0)',
  'code-37': '#,##0_);(#,##0)',
  'code-elapsed': '[h]:mm:ss',
}

const INTO = ['15.03.2026', '3/15/2026', '5%', '1 1/2', '=TODAY()', '=СЕГОДНЯ()', '1,5E+3', '1.5E+3']

export function buildEntryCases(): EntryCase[] {
  const { out, add } = collector()
  for (const [group, texts] of Object.entries(GENERAL)) for (const t of texts) add(group, t)
  for (const [name, format] of Object.entries(FORMATS)) for (const t of INTO) add(`in-${name}`, t, format)
  return out
}
