/** Checks shared by the tests of stored sets of Excel's answers. */
import { expect, it } from 'vitest'

/** A test that the set's stored cases are those its module builds now (`built`), and that each has an answer. */
export function currentCases<C extends { id: string }>(set: { cases: C[]; answers: Record<string, unknown> }, built: C[]): void {
  it('the stored cases are the current ones, all answered', () => {
    expect(set.cases).toEqual(built)
    expect(built.filter((c) => !set.answers[c.id])).toEqual([])
  })
}
