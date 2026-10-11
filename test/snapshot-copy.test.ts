import { describe, expect, it } from 'vitest'
import { Budget, CalcError, withBudget } from '../src/exec/budget.js'
import { snapshotBook } from '../src/session/snapshot.js'
import type { BookInput } from '../src/contract.js'

describe('snapshot copying can stop before a large sheet is complete', () => {
  it('yields while copying and honors cancellation without changing its input', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: Object.fromEntries(Array.from({ length: 4096 }, (_, i) => [`A${i + 1}`, { value: i }])) }] }
    const copy = snapshotBook(book)
    const budget = new Budget(undefined, Infinity, undefined, 0)
    budget.slice()
    expect(withBudget(budget, () => copy.next()).done).toBe(false)
    const stopped = new CalcError('aborted', 'snapshot copying cancelled')
    budget.halted = stopped
    expect(() => withBudget(budget, () => copy.next())).toThrow(stopped)
    expect(Object.keys(book.sheets[0].cells)).toHaveLength(4096)
    expect(book.sheets[0].cells.A4096.value).toBe(4095)
  })
})
