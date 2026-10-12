/**
 * The grid's own model of Find as a find set's predictions, written before Excel runs the set: `npx tsx
 * bench/coverage/find/predict.ts <set>` writes `find/<set>.predicted.json`, each case's cells in each of the package's
 * locales. The model is `xlcalc-sheet/src/find-match.ts`, which searched a session's windows before the session found
 * by Excel's rules: the query's wildcards as the lookup functions read them (`wildSource`), the whole text with Match
 * entire cell contents, any case without Match case; Look in Formulas the text a cell opens with (`editText` of
 * `getViewport`), Values the text it shows; an empty query finds nothing, a cell with an empty text is never found;
 * hidden rows count; the cells by rows, or by columns.
 */
import { fileURLToPath } from 'node:url'
import { createSession } from '../../../src/index.js'
import { wildSource } from '../../../src/criteria.js'
import { formatAddress, parseAddress } from '../../../src/model.js'
import { parseRange } from '../../../src/structure/area.js'
import { TAGS } from '../locales.js'
import { setModule, writeAnswers } from '../stored.js'
import { findBook, type FindBook, type FindQuery, type FindSet } from './cases.js'

const set = process.argv[2]
const { BOOKS, buildFindCases } = await setModule<FindSet>('find', set)
const clock = { now: '2026-10-12T12:00:00Z', seed: 1 }

function predict(b: FindBook, q: FindQuery, tag: string): string[] {
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

const cases = buildFindCases()
const locales = Object.fromEntries(TAGS.map((tag) => [tag, Object.fromEntries(cases.map((c) => [c.id, predict(BOOKS[c.book], c.query, tag)]))]))
writeAnswers(fileURLToPath(new URL(`./${set}.predicted.json`, import.meta.url)), { model: 'xlcalc-sheet find-match.ts', locales })
