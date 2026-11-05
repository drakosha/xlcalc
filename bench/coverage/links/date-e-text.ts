/**
 * TEXT with the date codes of `display/date-e-held.ts`, each case a file of its own, and an upper-case exponent beside them. Predicted
 * before Excel ran the set: `e` and its runs are `yyyy`, and a lower-case exponent stays what it is, `#VALUE!`.
 */
import type { CachedValue, LinksCase, LinksCases } from './cases.js'

const one = (id: string, code: string, predicted: CachedValue): LinksCase => ({
  id,
  about: `TEXT with ${code}`,
  sheets: [{ name: 'S', cells: { B1: { value: 1 } } }],
  linked: { S: { A1: `TEXT(46096.5,"${code}")` } },
  links: [],
  read: [{ sheet: 'S', range: 'A1' }],
  predicted: { 'S!A1': predicted },
})

export function buildLinksCases(): LinksCases {
  return {
    cases: [
      one('e', 'e', '2026'),
      one('ee', 'ee', '2026'),
      one('eee', 'eee', '2026'),
      one('upper', 'E', '2026'),
      one('ymd', 'e/m/d', '2026/3/15'),
      one('exponent', '0E+00', '5E+04'),
      one('lower-exponent', '0e+00', { error: 'VALUE' }),
      one('digit-e', '0e', { error: 'VALUE' }),
    ],
  }
}
