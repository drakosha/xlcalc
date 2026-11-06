/**
 * TEXT with words that hold an `e`, each case a file of its own, held out after `date-e-text`, whose rule took any `e` for a year.
 * Predicted before Excel ran the set: `e` is a year among date codes alone, `me` as `m` and a year; a word with any other letter is
 * no format, `#VALUE!`, as it was before `e` was a year; `General` is itself in en-US.
 */
import type { CachedValue, LinksCase, LinksCases } from './cases.js'

const VALUE: CachedValue = { error: 'VALUE' }

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
      ...['red', 'hello', 'seven', 'year', 'date', 'the', 'en', 'ne', 'Gene', 'ae', 're', 'le'].map((w) => one(w.toLowerCase() === w ? w : `upper-${w}`, w, VALUE)),
      one('general', 'General', '46096.5'),
      one('me', 'me', '32026'),
      one('em', 'em', '20263'),
      one('de', 'de', '152026'),
      one('ed', 'ed', '202615'),
      one('he', 'he', '122026'),
      one('ye', 'ye', '262026'),
      one('se', 'se', '02026'),
      one('e-am', 'e AM/PM', '2026 PM'),
    ],
  }
}
