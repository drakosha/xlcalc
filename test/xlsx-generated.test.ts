/** The xlsx subpath's generated tables are current: the `_xlfn.` prefixes from the harness's data, the argument modes from the engine's definitions. */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { MODES, modesSource, PREFIXES, prefixesSource } from '../bench/coverage/xlsx/generate.js'

describe('xlsx generated tables', () => {
  it('prefixes.ts and modes.ts are what bench/coverage/xlsx/generate.ts writes now', () => {
    expect(readFileSync(PREFIXES, 'utf8')).toBe(prefixesSource())
    expect(readFileSync(MODES, 'utf8')).toBe(modesSource())
  })
})
