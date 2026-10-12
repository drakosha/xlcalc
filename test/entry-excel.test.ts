/**
 * `parseEntry` against Excel for Mac's answers to the entry sets, en-US and ru-RU: the texts set through
 * `Formula2Local` (`typed`, `typed-held`, `typed-held-2`) and typed on the keyboard (`typed-keys`, `typed-keys-held`,
 * `typed-keys-held-2`, `typed-keys-held-3`). The held-out sets keep the answers the engine predicted before Excel ran them
 * (`bench/coverage/entry/<set>.predicted.json`). Known differences are in `test/entry-known.ts`.
 */
import { describe, expect, it } from 'vitest'
import { ENTRY_KNOWN, ENTRY_MAC_KNOWN } from './entry-known.js'
import { checkCase, entrySets } from './entry-sets.js'

const SETS = entrySets('')

describe.each(SETS)('entry $tag $name (Excel for Mac)', (set) => {
  it.each(set.cases.map((c) => [c.id, c] as const))('%s', (_, c) => checkCase(set, c, (k) => ENTRY_KNOWN.get(k) ?? ENTRY_MAC_KNOWN.get(k)))
})

describe('entry sets (Excel for Mac)', () => {
  it('every known difference is a stored case', () => {
    const stored = new Set(SETS.flatMap((s) => s.cases.map((c) => `${s.tag}:${s.name}:${c.id}`)))
    expect([...ENTRY_KNOWN.keys(), ...ENTRY_MAC_KNOWN.keys()].filter((k) => !stored.has(k))).toEqual([])
  })
})
