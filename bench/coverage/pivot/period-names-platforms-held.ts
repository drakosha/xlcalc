/** The same groupings made independently in Excel for Mac and Windows in ru-RU. */
import type { Period, PivotCase, PivotOp } from './cases.js'
import { serial } from './data.js'

type Group = Extract<PivotOp, { group: string }>

export function buildPivotCases(): PivotCase[] {
  function grouped(id: string, group: Group, before: PivotOp[] = []): PivotCase {
    const field = group.group === 'Qty' ? 'Amount' : 'Qty'
    return {
      id, data: 'stock',
      ops: [...before, { rows: [group.group] }, { value: { field, fn: 'count' } }, { layout: 'tabular' }, group],
    }
  }
  return [
    grouped('months', { group: 'Date', periods: ['months'] }),
    grouped('days', { group: 'Date', periods: ['days'] }),
    grouped('years-quarters-months', {
      group: 'Date', periods: ['years', 'quarters', 'months'],
      start: { d: serial(2024, 3, 15) }, end: { d: serial(2025, 8, 15) },
    }),
    grouped('fortnight', {
      group: 'Date', periods: ['days'], by: 14,
      start: { d: serial(2024, 3, 15) }, end: { d: serial(2025, 8, 15) },
    }),
    grouped('hours-minutes-seconds', { group: 'When', periods: ['hours', 'minutes', 'seconds'] }, [
      { edit: 'D!J2', to: serial(2025, 1, 1, 0, 5) + 7 / 86_400 },
      { edit: 'D!J3', to: serial(2025, 2, 2, 13, 21) + 49 / 86_400 },
      { refresh: true },
    ]),
    grouped('numeric-integer', { group: 'Qty', start: 2, end: 8, by: 2 }),
    grouped('numeric-fraction', { group: 'Qty', start: 2, end: 8, by: 1.5 }),
  ]
}
