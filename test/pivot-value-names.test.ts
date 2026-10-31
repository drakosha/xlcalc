/**
 * A value field's name when its function changes, and a calculated field added through the object model, against
 * Excel for Windows in en-US and ru-RU (the sets `value-names` and `value-names-reopen`, answers in
 * `test/excel/win/pivot/[ru-RU/]`): each case played on the engine as `bench/coverage/pivot/value-names-predict.ts`
 * plays it, its function changed through `editPivotTable`; the cells of the pivot table's area after each phase are
 * Excel's. Excel names the field anew whatever its name: the old function's words of the workbook's language before the
 * rest become the new function's, any other name (its own, another language's default) the new default name. The same
 * way the set `period-names` and its held-out `period-names-held`: the fields a grouping of dates makes and its groups' labels in the workbook's language,
 * as Excel's interface names them (`Years (Date)`, `Годы (Date)`; `Qtr1`, `Кв-л1`; `<1/1/2024`, `<01.01.2024`; `<10 or
 * (blank)`, `<10 или (пусто)`).
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { predictedNames } from '../bench/coverage/pivot/value-names-predict.js'
import type { PivotAnswer, PivotCase, PivotSet } from '../bench/coverage/pivot/cases.js'

const SETS = ['value-names', 'value-names-reopen', 'period-names', 'period-names-held']
/** Cases left out, with why: Excel's object model took the field off the rows as it became a value field, so the step grouped nothing (`AddDataField`). */
const LEFT_OUT: Record<string, string> = { 'period-names-held/qty-fraction-start': 'Qty off the rows as a value field' }
const LOCALES = [undefined, 'ru-RU']

describe.each(SETS.flatMap((set) => LOCALES.map((locale) => [set, locale] as const)))('the value fields of the set %s (%s)', (set, locale) => {
  const answers = JSON.parse(readFileSync(new URL(`excel/win/pivot/${locale ? `${locale}/` : ''}${set}.json`, import.meta.url), 'utf8')) as { answers: Record<string, PivotAnswer> }
  it('are named as Excel names them', async () => {
    const { buildPivotCases } = (await import(`../bench/coverage/pivot/${set}.ts`)) as PivotSet
    const misses: string[] = []
    for (const c of (buildPivotCases() as PivotCase[]).filter((x) => !LEFT_OUT[`${set}/${x.id}`])) {
      const want = answers.answers[c.id].phases.map((p) => {
        const a = p.pivots.PT1
        return 'error' in a ? a.error : a.values
      })
      const got = await predictedNames(c, locale)
      if (JSON.stringify(got) !== JSON.stringify(want)) misses.push(`${c.id}: ours ${JSON.stringify(got.map((x) => x[0]))}, Excel ${JSON.stringify(want.map((x) => (typeof x === 'string' ? x : x[0])))}`)
    }
    expect(misses).toEqual([])
  })
})
