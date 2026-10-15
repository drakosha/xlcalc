/**
 * `docs/api.md` is the reference `test/api-reference.ts` makes from the package's types: a change of an export or of its
 * comment shows there (`npx tsx test/api-reference.ts --write`), and every function, value and type has a comment.
 */
import { describe, expect, it } from 'vitest'
import { apiReference, current } from './api-reference.js'

describe('docs/api.md', () => {
  const made = apiReference()

  it('is the reference the package types make', () => {
    expect(current() === made, 'docs/api.md is out of date: npx tsx test/api-reference.ts --write').toBe(true)
  }, 60_000)

  it('has a comment for every export', () => {
    const bare = made
      .split(/\n(?=### )/)
      .filter((b) => b.startsWith('### '))
      .filter((b) => !b.slice(b.indexOf('\n')).replace(/```[\s\S]*?```/g, '').trim())
      .map((b) => b.slice(0, b.indexOf('\n')))
    expect(bare).toEqual([])
  })
})
