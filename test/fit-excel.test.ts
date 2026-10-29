/**
 * AutoFit and the width typing gives a column against Excel for Mac's answers to the fit sets, en-US and ru-RU
 * (`test/excel/fit/`): the width Excel saved for each case's column and what its cells show after, by `fitColumns`,
 * `parseEntry` and `entryWidth`. Known differences are in `test/fit-known.ts`. The engine's predictions, written before
 * Excel ran each set (`bench/coverage/fit/<set>.predicted.json`), are kept with how many Excel confirmed.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { FitPrediction } from '../bench/coverage/fit/engine.js'
import { FIT_KNOWN, FIT_REASONS } from './fit-known.js'
import { currentCases, engineAnswer, excelWidth, fitSets } from './fit-sets.js'

const SETS = fitSets('')

describe.each(SETS)('fit $tag $name (Excel for Mac)', (set) => {
  it.each(set.cases.map((c) => [c.id, c] as const))('%s', async (_, c) => {
    expect((await currentCases(set.name, set.tag)).find((x) => x.id === c.id), "the stored case is the generator's current one").toEqual(c)
    const a = set.answers[c.id]
    const e = engineAnswer(set, c)
    const got = { width: e.width, texts: e.texts }
    const excel = { width: excelWidth(a), texts: a.texts }
    const reason = FIT_KNOWN.get(`${set.tag}:${set.name}:${c.id}`)
    if (reason) expect(got, `${c.id} is known to differ: ${FIT_REASONS[reason]}`).not.toEqual(excel)
    else expect(got).toEqual(excel)
  })
})

/** How many of each set's predictions Excel for Mac confirmed (the width of the column), by `<tag>:<set>`. */
const CONFIRMED: Record<string, number> = {
  'en-US:autofit': 11,
  'ru-RU:autofit': 11,
  'en-US:typed': 21,
  'ru-RU:typed': 19,
  'en-US:glyphs': 253,
  'en-US:held': 76,
  'en-US:rotation': 14,
  'en-US:rotation-2': 3,
  'ru-RU:held': 74,
}

describe('fit sets (Excel for Mac)', () => {
  it('every known difference is a stored case', () => {
    const stored = new Set(SETS.flatMap((s) => s.cases.map((c) => `${s.tag}:${s.name}:${c.id}`)))
    expect([...FIT_KNOWN.keys()].filter((k) => !stored.has(k))).toEqual([])
  })
  it.each(SETS.map((s) => [`${s.tag}:${s.name}`, s] as const))('%s: the predictions written before Excel ran it', (key, set) => {
    const predicted = JSON.parse(readFileSync(new URL(`../bench/coverage/fit/${set.name}.predicted.json`, import.meta.url), 'utf8')).locales[set.tag] as Record<string, FitPrediction>
    expect(Object.keys(predicted)).toEqual(set.cases.map((c) => c.id))
    expect(set.cases.filter((c) => predicted[c.id].width === excelWidth(set.answers[c.id])).length).toBe(CONFIRMED[key])
  })
})
