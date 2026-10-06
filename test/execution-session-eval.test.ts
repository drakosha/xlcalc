/** Sessions across the execution modes (`execution-compare.ts`): an eval session equals a full closures computation (no code generation) after each of 200 seeded edits. */
import { describe, it } from 'vitest'
import { sessionAgainstFull } from './execution-compare.js'

describe('execution modes against each other', () => {
  it('200 session edits: an eval session equals a full closures computation (no code generation) after each one', () => sessionAgainstFull('eval'), 300_000)
})
