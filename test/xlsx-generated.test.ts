/** The xlsx subpath's generated table is current: the `_xlfn.` prefixes from the harness's data. */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { PREFIXES, prefixesSource } from '../bench/coverage/xlsx/generate.js'

describe('xlsx generated tables', () => {
  it('prefixes.ts is what bench/coverage/xlsx/generate.ts writes now', () => {
    expect(readFileSync(PREFIXES, 'utf8')).toBe(prefixesSource())
  })
})
