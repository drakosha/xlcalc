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
import type { CellValue } from '../src/contract.js'
import { buildPivotCases as platformCases } from '../bench/coverage/pivot/period-names-platforms-held.js'
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

/** Excel for Mac's ru-RU labels of months, including within the labels of groups by day of year. */
const MAC_MONTHS: Record<string, string> = { фев: 'февр', мар: 'март', июн: 'июнь', июл: 'июль', сен: 'сент', ноя: 'нояб' }
const MONTH_COLUMNS: Record<string, number> = { months: 0, days: 0, 'years-quarters-months': 2 }
const PLATFORM_SET = 'period-names-platforms-held'
const PLATFORM_LOCALE = 'ru-RU'
const macGroups = JSON.parse(readFileSync(new URL(`excel/pivot/${PLATFORM_LOCALE}/${PLATFORM_SET}.json`, import.meta.url), 'utf8')) as {
  answers: Record<string, { values: CellValue[][] }>
}

describe('Russian labels of range groups in Excel for Mac', () => {
  it.each(platformCases())('$id has the recorded Mac differences', async (c) => {
    const [area] = await predictedNames(c, PLATFORM_LOCALE)
    const macArea = area.map((row) => row.map((value, column) => {
      if (typeof value !== 'string') return value
      if (column === MONTH_COLUMNS[c.id]) {
        return value.replace(/^(\d{2}\.)?(фев|мар|июн|июл|сен|ноя)$/, (_, prefix = '', month: string) => prefix + MAC_MONTHS[month])
      }
      if (c.id === 'hours-minutes-seconds' && column === 0) return value.replace(/^\d( Итог)?$/, '0$&')
      return value
    }))
    expect(macArea).toEqual(macGroups.answers[c.id].values)
  })
})
