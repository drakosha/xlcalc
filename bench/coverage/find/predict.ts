/**
 * A find set's predictions, written before Excel runs the set: `npx tsx bench/coverage/find/predict.ts <set> [grid]`
 * writes `find/<set>.predicted.json`, each case's cells in each of the package's locales. By default they are the
 * engine's of the day (`findInBook`: `findCells` with Find's options), so a set's committed predictions stay what the
 * engine answered before Excel ran it. With `grid`, the model the grid searched by
 * before the session found by Excel's rules (`xlcalc-sheet/src/find-match.ts`, for `rules-held`): the query's wildcards
 * as the lookup functions read them (`wildSource`), Look in Formulas `editText` of `getViewport`, Values the text it
 * shows, hidden rows found, an empty query finding nothing, a cell with an empty text never found.
 */
import { createSession } from '../../../src/index.js'
import { wildSource } from '../../../src/criteria.js'
import { formatAddress, parseAddress } from '../../../src/model.js'
import { parseRange } from '../../../src/structure/area.js'
import { TAGS } from '../locales.js'
import { setModule, writePrediction } from '../stored.js'
import { findBook, findInBook, type FindBook, type FindQuery, type FindSet } from './cases.js'

const [set, model] = process.argv.slice(2)
const { BOOKS, buildFindCases } = await setModule<FindSet>('find', set)
const clock = { now: '2026-10-12T12:00:00Z', seed: 1 }

/** The grid's model of a case. */
function grid(b: FindBook, q: FindQuery, tag: string): string[] {
  if (!q.what) return []
  const re = new RegExp(q.whole ? `^(?:${wildSource(q.what)})$` : wildSource(q.what), q.matchCase ? 'u' : 'iu')
  const s = createSession(findBook(b, tag), { clock })
  try {
    const used = s.getUsedRange('S')
    const v = s.getViewport('S', `A1:${formatAddress(used.rows - 1, used.columns - 1)}`, { editText: true })
    const x = q.range ? parseRange(q.range)! : null
    const hits = Object.entries(v.cells).flatMap(([cell, c]) => {
      const p = parseAddress(cell)
      const text = (q.lookIn === 'formulas' ? c.editText : c.text) ?? ''
      return text !== '' && (!x || (p.row >= x.r1 && p.row <= x.r2 && p.col >= x.c1 && p.col <= x.c2)) && re.test(text) ? [{ cell, ...p }] : []
    })
    return hits.sort(q.byColumns ? (a, b) => a.col - b.col || a.row - b.row : (a, b) => a.row - b.row || a.col - b.col).map((h) => h.cell)
  } finally {
    s.close()
  }
}

const predict = model === 'grid' ? grid : findInBook
const cases = buildFindCases()
const locales = Object.fromEntries(TAGS.map((tag) => [tag, Object.fromEntries(cases.map((c) => [c.id, predict(BOOKS[c.book], c.query, tag)]))]))
writePrediction(import.meta.url, `${set}`, { model: model === 'grid' ? 'xlcalc-sheet find-match.ts' : 'findCells', locales })
