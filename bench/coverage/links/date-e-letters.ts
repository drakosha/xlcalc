/**
 * TEXT with `e` beside each other letter, each case a file of its own, held out after `date-e-words`, where `n` refused and `r`, `l`,
 * `a`, `t` and `o` did not. Predicted before Excel ran the set: only `n` refuses, a date's letter is its code, any other is itself.
 */
import type { CachedValue, LinksCase, LinksCases } from './cases.js'

const CODES: Record<string, string> = { d: '15', h: '12', m: '3', s: '0', y: '26' }
const VALUE: CachedValue = { error: 'VALUE' }
const YEAR = '2026'

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
  const cases: LinksCase[] = []
  for (const c of 'abcdfghijklmnopqrstuvwxyz') {
    const self = CODES[c] ?? c
    cases.push(one(`e-${c}`, `e${c}`, c === 'n' ? VALUE : `${YEAR}${self}`))
    cases.push(one(`${c}-e`, `${c}e`, c === 'n' ? VALUE : `${self}${YEAR}`))
  }
  return { cases }
}
