/**
 * Row heights and wrapped entries against Excel's answers to the rows sets (`test/excel/rows/`, `test/excel/win/rows/`):
 * each case's book in a session, its texts typed, entered into a selection or pasted and its row fitted, as
 * `bench/coverage/rows/engine.ts` does it. Against Excel for Mac the row's height and each cell's content and wrapping;
 * against Excel for Windows, which fits rows to its pixels (a line of Calibri 11 is 20 pixels, 15 points, wrapped or
 * not), the cells alone and how many heights are the Mac's. Known differences are in `test/rows-known.ts`. The engine's
 * predictions, written before Excel ran each set (`bench/coverage/rows/<set>.predicted.json`), are kept with how many
 * heights Excel for Mac confirmed.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredRows } from '../bench/coverage/rows.js'
import type { RowCase, RowSet } from '../bench/coverage/rows/cases.js'
import { engineRows, type RowsPrediction } from '../bench/coverage/rows/engine.js'
import { createSession } from '../src/index.js'
import { ROWS_KNOWN, ROWS_REASONS, ROWS_WIN_KNOWN } from './rows-known.js'
import { localeSets } from './stored-sets.js'
import './locales.js'

interface RowSetFile extends StoredRows {
  name: string
  tag: string
}

/** The sets stored under `test/excel/<root>rows/`, every locale's. */
function rowSets(root: '' | 'win/'): RowSetFile[] {
  return localeSets<StoredRows>(`${root}rows`).map((r) => ({ name: r.set, tag: r.tag, ...r.stored }))
}

const engine = (set: RowSetFile, c: RowCase) => engineRows({ createSession }, c, set.tag, String(set.meta.date))

const MAC = rowSets('')
const WIN = rowSets('win/')

describe.each(MAC)('rows $tag $name (Excel for Mac)', (set) => {
  it.each(set.cases.map((c) => [c.id, c] as const))('%s', async (_, c) => {
    const now = ((await import(`../bench/coverage/rows/${set.name}.ts`)) as RowSet).buildRowCases(set.tag)
    expect(now.find((x) => x.id === c.id), "the stored case is the generator's current one").toEqual(c)
    const a = set.answers[c.id]
    const e = engine(set, c)
    const got = { height: e.height, cells: e.cells, more: e.more }
    const excel = { height: a.height, cells: a.cells, more: a.more }
    const reason = ROWS_KNOWN.get(`${set.tag}:${set.name}:${c.id}`)
    if (reason) expect(got, `${c.id} is known to differ: ${ROWS_REASONS[reason]}`).not.toEqual(excel)
    else expect(got).toEqual(excel)
  })
})

describe.each(WIN)('rows $tag $name (Excel for Windows)', (set) => {
  it.each(set.cases.map((c) => [c.id, c] as const))('%s: the cells as the engine leaves them', (_, c) => {
    const e = engine(set, c)
    const reason = ROWS_WIN_KNOWN.get(`${set.tag}:${set.name}:${c.id}`)
    if (reason) expect(e.cells, `${c.id} is known to differ: ${ROWS_REASONS[reason]}`).not.toEqual(set.answers[c.id].cells)
    else expect(e.cells).toEqual(set.answers[c.id].cells)
  })
})

/** How many of each set's predictions Excel for Mac confirmed (the row's height), by `<tag>:<set>`. */
const CONFIRMED: Record<string, number> = {
  'en-US:autofit': 50,
  'ru-RU:autofit': 50,
  'en-US:held': 81,
  'ru-RU:held': 81,
  'en-US:lines': 79,
  'ru-RU:lines': 79,
  'en-US:keys': 19,
  'en-US:quotes': 12,
  'en-US:rotation': 7,
  'en-US:rotation-2': 0,
  'en-US:patched': 4,
}

/** The cases whose height Excel for Windows gave as the Mac did, by `<tag>:<set>`. */
const WIN_SAME_HEIGHT: Record<string, number> = {
  'en-US:autofit': 49,
  'ru-RU:autofit': 49,
  'en-US:held': 15,
  'en-US:keys': 6,
  'en-US:lines': 7,
  'en-US:quotes': 12,
  'en-US:rotation': 14,
  'en-US:rotation-2': 8,
}

describe('rows sets', () => {
  it('every known difference is a stored case', () => {
    const stored = new Set([...MAC, ...WIN].flatMap((s) => s.cases.map((c) => `${s.tag}:${s.name}:${c.id}`)))
    expect([...ROWS_KNOWN.keys(), ...ROWS_WIN_KNOWN.keys()].filter((k) => !stored.has(k))).toEqual([])
  })
  it.each(MAC.map((s) => [`${s.tag}:${s.name}`, s] as const))('%s: the predictions written before Excel ran it', (key, set) => {
    const predicted = JSON.parse(readFileSync(new URL(`../bench/coverage/rows/${set.name}.predicted.json`, import.meta.url), 'utf8')).locales[set.tag] as Record<string, RowsPrediction>
    expect(Object.keys(predicted)).toEqual(set.cases.map((c) => c.id))
    expect(set.cases.filter((c) => predicted[c.id].height === set.answers[c.id].height).length).toBe(CONFIRMED[key])
  })
  it.each(WIN.map((s) => [`${s.tag}:${s.name}`, s] as const))('%s: the heights Excel for Windows gives as the Mac', (key, set) => {
    const mac = MAC.find((s) => s.tag === set.tag && s.name === set.name)
    if (!mac) return
    expect(set.cases.filter((c) => mac.answers[c.id] && set.answers[c.id].height === mac.answers[c.id].height).length).toBe(WIN_SAME_HEIGHT[key])
  })
})
