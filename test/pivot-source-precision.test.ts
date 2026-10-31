import { expect, it } from 'vitest'
import { createSession, registerPivot } from '../src/index.js'
import { pivotEngine } from '../src/pivot/index.js'

it('refreshes from formula values before the public result rounds them', () => {
  registerPivot(pivotEngine)
  const s = createSession({ sheets: [
    { name: 'D', cells: { A1: { value: 'R' }, B1: { value: 'V' }, A2: { value: 'A' }, B2: { formula: '801/43' } } },
    { name: 'P', cells: {} },
  ] })
  expect(s.editStructure({ op: 'addPivotTable', sheet: 'P', cell: 'A3', source: { type: 'range', sheet: 'D', range: 'A1:B2' }, definition: { rows: [0], values: [{ field: 1 }] } }).ok).toBe(true)
  expect(s.getBook().pivotCaches![0].records![0][1]).toBe(18.627906976744185)
  expect(s.getValues('D', 'B2')).toEqual([[18.6279069767442]])
  s.close()
})
