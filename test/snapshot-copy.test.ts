import { describe, expect, it } from 'vitest'
import { Budget, CalcError, run, withBudget } from '../src/exec/budget.js'
import { snapshotBook } from '../src/session/snapshot.js'
import type { BookInput } from '../src/contract.js'

describe('snapshot copying can stop before a large sheet is complete', () => {
  it.each([0, 1, 1023, 1024, 1025, 2049])('copies all %i cells and the last partial portion, without sharing nested objects', (size) => {
    const book: BookInput = {
      styles: [{ font: { bold: true } }],
      sheets: [{ name: 'S', cells: Object.fromEntries(Array.from({ length: size }, (_, i) => [`A${i + 1}`, { value: i, meta: { nested: [{ id: i }] }, note: { text: String(i) } }])) }, { name: 'Empty', cells: {} }],
    }
    const copied = run(snapshotBook(book))
    expect(copied).toEqual(book)
    copied.styles![0].font!.bold = false
    expect(book.styles![0].font!.bold).toBe(true)
    if (size) {
      const addr = `A${size}`
      const meta = copied.sheets[0].cells[addr].meta as { nested: { id: number }[] }
      meta.nested[0].id = -1
      copied.sheets[0].cells[addr].note!.text = 'copy only'
      expect(book.sheets[0].cells[addr]).toEqual({ value: size - 1, meta: { nested: [{ id: size - 1 }] }, note: { text: String(size - 1) } })
    }
  })

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
