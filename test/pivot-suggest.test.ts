/**
 * `suggestPivotField`: where Excel's PivotTable Fields list puts a field checked there, by the kinds of value the
 * cache holds, from its facts, its items or its rows.
 */
import { describe, expect, it } from 'vitest'
import type { PivotCache, PivotCacheField } from '../src/contract.js'
import { InputError, suggestPivotField } from '../src/index.js'

const cache = (fields: PivotCacheField[], records?: PivotCache['records']): PivotCache => ({ id: 1, source: { type: 'range', sheet: 'D', range: 'A1:C9' }, fields, ...(records ? { records } : {}) })
const by = (types: NonNullable<PivotCacheField['facts']>['types']) => suggestPivotField(cache([{ name: 'F', facts: { types } }]), 0)

describe('suggestPivotField', () => {
  it('numbers to Values with Sum, with empty cells with Count; the rest to the rows', () => {
    expect(by(['number'])).toEqual({ axis: 'values', function: 'sum' })
    expect(by(['number', 'blank'])).toEqual({ axis: 'values', function: 'count' })
    for (const types of [['string'], ['date'], ['boolean'], ['blank'], [], ['number', 'string'], ['number', 'date'], ['number', 'error'], ['number', 'boolean']] as const) expect(by([...types]), types.join()).toEqual({ axis: 'rows' })
  })

  it('takes the kinds from the items or the rows without facts; a calculated field to Values, groups to the rows', () => {
    expect(suggestPivotField(cache([{ name: 'R', items: ['a', null] }, { name: 'N' }], [[0, 1], [1, null]]), 1)).toEqual({ axis: 'values', function: 'count' })
    expect(suggestPivotField(cache([{ name: 'N', items: [1, 2] }]), 0)).toEqual({ axis: 'values', function: 'sum' })
    expect(suggestPivotField(cache([{ name: 'N' }]), 0)).toEqual({ axis: 'rows' })
    expect(suggestPivotField(cache([{ name: 'T', formula: 'N*2' }]), 0)).toEqual({ axis: 'values', function: 'sum' })
    expect(suggestPivotField(cache([{ name: 'N', facts: { types: ['number'] } }, { name: 'N2', group: { base: 0, range: { by: 'numbers', interval: 10 } } }]), 1)).toEqual({ axis: 'rows' })
  })

  it('wrong input is an InputError', () => {
    expect(() => suggestPivotField(cache([{ name: 'N' }]), 1)).toThrow(new InputError('field', 'expected an integer from 0 to 0, got 1'))
    expect(() => suggestPivotField(null as never, 0)).toThrow(InputError)
    expect(() => suggestPivotField(cache([]), 0)).toThrow(InputError)
  })
})
