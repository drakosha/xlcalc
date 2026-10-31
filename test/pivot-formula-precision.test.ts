import { expect, it } from 'vitest'
import { calculate } from '../src/pivot/formulas.js'

it('keeps the calculated-field result at binary64 precision for later calculations', () => {
  const [ratio] = calculate([{ text: '__pivot0_/__pivot1_', values: [801, 43] }])
  expect(ratio).toBe(18.627906976744185)
  expect(calculate([{ text: '__pivot0_*__pivot1_', values: [ratio, 43] }])).toEqual([801])
})
