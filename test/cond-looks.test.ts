/**
 * What conditional formats give cells, rule by rule, by the rules measured in Excel (`bench/coverage/cond/`, the sets
 * looks and looks-ru): comparisons as `=` and `<` compare with a blank operand as 0, formulas at each cell from the top-left
 * of the rule's cells, IF's reading of a formula's answer, blank rules by the value, statistics without errors and text,
 * duplicates by Excel's comparison, colour channels cut toward the start colour, data bar geometry, icons, priorities.
 */
import { describe, expect, it } from 'vitest'
import type { BookInput, CellInput, ConditionalFormat, ConditionalLook } from '../src/contract.js'
import { createSession } from '../src/index.js'
import { barOf, between } from '../src/cond/looks.js'
import './locales.js'

const CLOCK = { now: '2026-10-10T12:00:00Z', seed: 1 }
const RED = { fill: { color: '#FFC7CE' } }

/** The looks of a column of values under rules, by address (cells without a look left out). */
function looks(values: (CellInput | undefined)[], rules: ConditionalFormat[], execution: 'closures' | 'eval', more: BookInput['sheets'][number]['cells'] = {}, locale?: string): Record<string, ConditionalLook> {
  const cells: Record<string, CellInput> = { ...more }
  values.forEach((v, i) => v && (cells[`A${i + 1}`] = v))
  const s = createSession({ sheets: [{ name: 'S', cells, conditionalFormats: rules }], ...(locale ? { locale } : {}) }, { clock: CLOCK, execution })
  const w = s.getCells('S', `A1:C${values.length}`)
  return Object.fromEntries(Object.entries(w).flatMap(([a, v]) => (v.appearance ? [[a, v.appearance]] : [])))
}
const held = (l: Record<string, ConditionalLook>) => Object.keys(l)
const v = (value: CellInput['value']): CellInput => ({ value })
const col = (n: number) => [`A1:A${n}`]

describe.each(['closures', 'eval'] as const)('looks (%s)', (execution) => {
  it('cell value rules compare as the operators do; a blank cell an operand refers to is 0', () => {
    const vals = [v(1), v(5), v('abc'), v(''), v(true), { formula: 'NA()' }, undefined, v(-3)]
    const r = (operator: Extract<ConditionalFormat, { type: 'cellIs' }>['operator'], formulas: string[]) => held(looks(vals, [{ type: 'cellIs', ranges: col(8), operator, formulas, style: RED }], execution))
    expect(r('greaterThan', ['3'])).toEqual(['A2', 'A3', 'A4', 'A5'])
    expect(r('lessThan', ['3'])).toEqual(['A1', 'A7', 'A8'])
    expect(r('between', ['10', '1'])).toEqual(['A1', 'A2'])
    expect(r('notBetween', ['1', '5'])).toEqual(['A3', 'A4', 'A5', 'A7', 'A8'])
    expect(r('equal', ['"ABC"'])).toEqual(['A3'])
    // Z1 is blank: as a formula's answer it is 0, so empty text (greater than numbers) is not at most it
    expect(r('lessThanOrEqual', ['$Z$1'])).toEqual(['A7', 'A8'])
  })

  it('a formula is written for the top-left of the rule cells and moves to each; its text TRUE holds; a range is not intersected', () => {
    const vals = [0, 1, 2, 3, 4, 5].map(v)
    expect(held(looks(vals, [{ type: 'expression', ranges: ['B3:B6', 'C1:C3'], formula: '$A3>2', style: RED }], execution))).toEqual(['C2', 'B3', 'C3', 'B4'])
    expect(held(looks(vals, [{ type: 'expression', ranges: col(6), formula: 'IF($A1>3,"TRUE","")', style: RED }], execution))).toEqual(['A5', 'A6'])
    expect(held(looks(vals, [{ type: 'expression', ranges: ['B1:B6'], formula: '$A$1:$A$6>3', style: RED }], execution))).toEqual([])
    expect(held(looks(vals, [{ type: 'expression', ranges: col(6), formula: '1/($A1-3)', style: RED }], execution))).toEqual(['A1', 'A2', 'A3', 'A5', 'A6'])
  })

  it('text rules compute their formula: wildcards, digits of numbers, any case; Excel computes the formula, not the text', () => {
    const vals = [v('Apple'), v('PINEAPPLE'), v('banana'), v(123), v('a*b'), v('axb')]
    const t = (text: string, type: 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith' = 'containsText', formula?: string) => held(looks(vals, [{ type, ranges: col(6), text, ...(formula ? { formula } : {}), style: RED }], execution))
    expect(t('apple')).toEqual(['A1', 'A2'])
    expect(t('apple', 'notContainsText')).toEqual(['A3', 'A4', 'A5', 'A6'])
    expect(t('ap', 'beginsWith')).toEqual(['A1'])
    expect(t('LE', 'endsWith')).toEqual(['A1', 'A2'])
    expect(t('a*b')).toEqual(['A5', 'A6'])
    expect(t('2')).toEqual(['A4'])
    expect(t('zzz', 'containsText', 'NOT(ISERROR(SEARCH("apple",A1)))')).toEqual(['A1', 'A2'])
  })

  it('blank rules go by the value: spaces are blank, an error is not', () => {
    const vals = [v(''), v(' '), undefined, v('x'), v(0), { formula: '""' }, { formula: 'NA()' }]
    expect(held(looks(vals, [{ type: 'containsBlanks', ranges: col(7), style: RED }], execution))).toEqual(['A1', 'A2', 'A3', 'A6'])
    expect(held(looks(vals, [{ type: 'notContainsBlanks', ranges: col(7), style: RED }], execution))).toEqual(['A4', 'A5', 'A7'])
    expect(held(looks(vals, [{ type: 'containsErrors', ranges: col(7), style: RED }], execution))).toEqual(['A7'])
  })

  it('time periods by the clock', () => {
    const vals = ['TODAY()', 'TODAY()-1', 'TODAY()+1', 'TODAY()-6', 'TODAY()-7', 'EDATE(TODAY(),-1)', 'EDATE(TODAY(),1)'].map((formula) => ({ formula }))
    const p = (period: Extract<ConditionalFormat, { type: 'timePeriod' }>['period']) => held(looks(vals, [{ type: 'timePeriod', ranges: col(7), period, style: RED }], execution))
    // 10 October 2026 is a Saturday
    expect(p('yesterday')).toEqual(['A2'])
    expect(p('last7Days')).toEqual(['A1', 'A2', 'A4'])
    expect(p('thisWeek')).toEqual(['A1', 'A2', 'A4'])
    expect(p('lastWeek')).toEqual(['A5'])
    expect(p('lastMonth')).toEqual(['A6'])
    expect(p('nextMonth')).toEqual(['A7'])
  })

  it('top, bottom, average and duplicates over the numbers, errors and text left out', () => {
    const nums = [10, 20, 20, 30, 40, 50, 60, 70, 80, 90, 100, 30].map(v)
    const r = (rule: Partial<ConditionalFormat>, vals: (CellInput | undefined)[] = nums) => held(looks(vals, [{ ranges: col(vals.length), style: RED, ...rule } as ConditionalFormat], execution))
    expect(r({ type: 'top10', rank: 3 })).toEqual(['A9', 'A10', 'A11'])
    expect(r({ type: 'top10', rank: 2, bottom: true })).toEqual(['A1', 'A2', 'A3'])
    // 10 percent of 12 numbers is one of them, 25 percent three
    expect(r({ type: 'top10', rank: 10, percent: true })).toEqual(['A11'])
    expect(r({ type: 'top10', rank: 25, percent: true })).toEqual(['A9', 'A10', 'A11'])
    expect(r({ type: 'top10', rank: 2 }, [v(1), { formula: 'NA()' }, v(5), v('9'), v(3)])).toEqual(['A3', 'A5'])
    expect(r({ type: 'aboveAverage', equal: true })).toEqual(['A6', 'A7', 'A8', 'A9', 'A10', 'A11'])
    // the population's deviation: 16.4 is past the mean of 1..9, 30 and itself by more than it, less than the sample's
    expect(r({ type: 'aboveAverage', stdDev: 1 }, [1, 2, 3, 4, 5, 6, 7, 8, 9, 30, 16.4].map(v))).toEqual(['A10', 'A11'])
    expect(r({ type: 'duplicateValues' }, [v(1), v('1'), v('a'), v('A'), v(true), v(2), undefined, v('')])).toEqual(['A1', 'A2', 'A3', 'A4', 'A8'])
    expect(r({ type: 'uniqueValues' }, [v(1), v('1'), v('a'), v('A'), v(true), v(2), undefined, v('')])).toEqual(['A5', 'A6'])
  })

  it('colour scales: each channel moves by its share cut toward the start colour; equal numbers get the last colour; errors left out', () => {
    const scale = (points: Extract<ConditionalFormat, { type: 'colorScale' }>['points'], vals: (CellInput | undefined)[]) =>
      Object.fromEntries(Object.entries(looks(vals, [{ type: 'colorScale', ranges: col(vals.length), points }], execution)).map(([a, l]) => [a, l.style!.fill!.color]))
    const two = [{ value: { type: 'min' as const }, color: '#F8696B' }, { value: { type: 'max' as const }, color: '#63BE7B' }]
    expect(scale(two, [1, 2, 10].map(v))).toEqual({ A1: '#F8696B', A2: '#E8726C', A3: '#63BE7B' })
    expect(scale(two, [v(5), v(5)])).toEqual({ A1: '#63BE7B', A2: '#63BE7B' })
    expect(scale(two, [v(1), { formula: 'NA()' }, v(2), v('x'), v(10)])).toEqual({ A1: '#F8696B', A3: '#E8726C', A5: '#63BE7B' })
    const three = [{ value: { type: 'min' as const }, color: '#F8696B' }, { value: { type: 'percentile' as const, value: '50' }, color: '#FFEB84' }, { value: { type: 'max' as const }, color: '#63BE7B' }]
    expect(scale(three, [1, 2, 3, 4, 5, 6, 7, 8, 100, 200].map(v)).A2).toBe('#F98570')
    expect(between('#F8696B', '#63BE7B', 1 / 9)).toBe('#E8726C')
  })

  it('stopIfTrue leaves out the rules after; earlier rules win field by field; a colour scale after a fill gives none; a number format shows the text', () => {
    const vals = [2, 4, 6].map(v)
    const gt = (x: number, style: object, stopIfTrue = false): ConditionalFormat => ({ type: 'cellIs', ranges: col(3), operator: 'greaterThan', formulas: [String(x)], style, ...(stopIfTrue ? { stopIfTrue } : {}) })
    expect(looks(vals, [gt(5, { fill: { color: '#FF0000' }, font: { bold: true } }), gt(3, { fill: { color: '#00FF00' }, font: { italic: true } })], execution)).toEqual({
      A2: { style: { fill: { color: '#00FF00' }, font: { italic: true } } },
      A3: { style: { fill: { color: '#FF0000' }, font: { bold: true, italic: true } } },
    })
    expect(looks(vals, [gt(5, { fill: { color: '#FF0000' } }, true), gt(3, { font: { italic: true } })], execution)).toEqual({ A2: { style: { font: { italic: true } } }, A3: { style: { fill: { color: '#FF0000' } } } })
    const scale: ConditionalFormat = { type: 'colorScale', ranges: col(3), points: [{ value: { type: 'min' }, color: '#000000' }, { value: { type: 'max' }, color: '#FFFFFF' }] }
    expect(looks(vals, [gt(5, { fill: { color: '#FF0000' } }), scale], execution).A3).toEqual({ style: { fill: { color: '#FF0000' } } })
    const s = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 6 } }, conditionalFormats: [gt(5, { numberFormat: '0.00%' })] }] }, { clock: CLOCK, execution })
    expect(s.getCells('S', 'A1').A1).toEqual({ value: 6, text: '600.00%', appearance: { style: { numberFormat: '0.00%' } } })
  })

  it('data bars and icon sets', () => {
    const vals = [0, 5, 10].map(v)
    const bars = looks(vals, [{ type: 'dataBar', ranges: col(3), min: { type: 'autoMin' }, max: { type: 'autoMax' }, color: '#638EC6', showValue: false }], execution)
    expect(bars.A2).toEqual({ bar: { start: 0, end: 0.5, color: '#638EC6', gradient: true }, hideValue: true })
    expect(bars.A1.bar?.end).toBe(0)
    const icons = (rule: Partial<Extract<ConditionalFormat, { type: 'iconSet' }>>) =>
      Object.fromEntries(
        Object.entries(looks(vals, [{ type: 'iconSet', ranges: col(3), icons: '3Arrows', points: [{ value: { type: 'percent', value: '0' } }, { value: { type: 'percent', value: '33' } }, { value: { type: 'percent', value: '67' } }], ...rule }], execution)).map(([a, l]) => [a, l.icon]),
      )
    expect(icons({})).toEqual({ A1: { set: '3Arrows', index: 0 }, A2: { set: '3Arrows', index: 1 }, A3: { set: '3Arrows', index: 2 } })
    expect(icons({ reverse: true }).A1).toEqual({ set: '3Arrows', index: 2 })
    expect(icons({ points: [{ value: { type: 'num', value: '0' }, icon: null }, { value: { type: 'num', value: '4' }, gte: false }, { value: { type: 'num', value: '10' }, icon: { set: '3Flags', index: 0 } }] })).toEqual({ A2: { set: '3Arrows', index: 1 }, A3: { set: '3Flags', index: 0 } })
  })

  it('in ru-RU a number reads with a decimal comma in a text rule', () => {
    expect(held(looks([v(1.5), v(2)], [{ type: 'containsText', ranges: col(2), text: ',', style: RED }], execution, {}, 'ru-RU'))).toEqual(['A1'])
  })

  it('a rectangle larger than 65,536 cells gives the looks of the cells it lists only', () => {
    const s = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, B5: { value: 2 } }, conditionalFormats: [{ type: 'containsBlanks', ranges: ['A:C'], style: RED }] }] }, { clock: CLOCK, execution })
    expect(Object.keys(s.getCells('S', 'A:C'))).toEqual(['A1', 'B5'])
    expect(Object.keys(s.getCells('S', 'A1:C2'))).toEqual(['A1', 'B1', 'C1', 'A2', 'B2', 'C2'])
  })
})

describe('data bar geometry', () => {
  const rule = { type: 'dataBar', ranges: ['A1'], min: { type: 'min' }, max: { type: 'max' }, color: '#000000' } as Extract<ConditionalFormat, { type: 'dataBar' }>
  it("Excel 2010's from the axis at the numbers' zero, Excel 2007's from the edge with lengths 10 to 90", () => {
    expect(barOf(rule, 5, 0, 10)).toEqual({ start: 0, end: 0.5, color: '#000000', gradient: true })
    expect(barOf({ ...rule, excel2007: true }, 0, 0, 10).end).toBeCloseTo(0.1)
    const neg = barOf({ ...rule, negativeColor: '#FF0000' }, -5, -10, 10)
    expect(neg).toEqual({ start: 0.25, end: 0.5, color: '#FF0000', gradient: true, axis: 0.5, axisColor: '#000000' })
    expect(barOf({ ...rule, axis: 'middle' }, 10, 0, 10)).toMatchObject({ start: 0.5, end: 1, axis: 0.5 })
  })
})
