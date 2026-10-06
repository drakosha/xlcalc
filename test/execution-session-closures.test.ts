/** Sessions across the execution modes (`execution-compare.ts`): a closures session (no code generation) equals a full eval computation after each of 200 seeded edits. */
import { describe, it } from 'vitest'
import { sessionAgainstFull } from './execution-compare.js'

describe('execution modes against each other', () => {
  it('200 session edits: a closures session (no code generation) equals a full eval computation after each one', () => sessionAgainstFull('closures'), 300_000)
})
