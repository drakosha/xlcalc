/**
 * `parseEntry` against Excel for Windows's answers to the entry sets set through `Formula2Local`, en-US and ru-RU
 * (`test/excel/win/entry/`). Excel for Windows is the reference where the two Excels part (`ENTRY_MAC_KNOWN`); the
 * known differences both share are in `test/entry-known.ts`.
 */
import { describe, expect, it } from 'vitest'
import { ENTRY_KNOWN, ENTRY_WIN_KNOWN } from './entry-known.js'
import { checkCase, entrySets } from './entry-sets.js'

describe.each(entrySets('win/'))('entry $tag $name (Excel for Windows)', (set) => {
  it.each(set.cases.map((c) => [c.id, c] as const))('%s', (_, c) => checkCase(set, c, (k) => ENTRY_KNOWN.get(k) ?? ENTRY_WIN_KNOWN.get(k)))
})

describe('entry sets (Excel for Windows)', () => {
  it('every known difference of Windows alone is a stored case', () => {
    const stored = new Set(entrySets('win/').flatMap((s) => s.cases.map((c) => `${s.tag}:${s.name}:${c.id}`)))
    expect([...ENTRY_WIN_KNOWN.keys()].filter((k) => !stored.has(k))).toEqual([])
  })
})
