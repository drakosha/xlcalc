/** Find and Replace over a host's texts (`findMatcher`, `findReplacer` of `xlcalc/text`) answer as a session's `findCells` and `replaceInCells` over the same texts. */
import { expect, it } from 'vitest'
import { createSession } from '../src/index.js'
import { findMatcher, findReplacer } from '../src/text.js'

const TEXTS = ['Straße', 'STRASSE', 'a*b', 'a1b2b', 'Σίσυφος', 'x~y', '']
const QUERIES: [string, object][] = [['straße', {}], ['STRA?E', { wildcards: true }], ['a~*b', { wildcards: true }], ['a*b', { wildcards: true, wholeCell: true }], ['σ', {}], ['Straße', { caseSensitive: true }], ['~', { wildcards: true }], ['', {}]]

it("matches and replaces as the session's Find and Replace", () => {
  const cells = Object.fromEntries(TEXTS.map((t, i) => [`A${i + 1}`, { value: t }]))
  const s = createSession({ sheets: [{ name: 'S', cells }] })
  for (const [q, o] of QUERIES) {
    const test = findMatcher(q, o)
    const found = s.findCells([{ sheet: 'S', range: 'A1:A7' }], q, { ...o, in: 'edit' }).cells.map((f) => TEXTS[Number(f.cell.slice(1)) - 1])
    expect([q, TEXTS.filter((t) => test?.(t))], JSON.stringify(o)).toEqual([q, found])
  }
  expect(findReplacer('a*b', 'X', { wildcards: true })?.('a1b2b')).toBe('X2b')
  expect(findReplacer('b', 'X')?.('abc')).toBe('aXc')
  expect(findReplacer('', 'X')).toBeNull()
  expect(() => findMatcher('a', { wholecell: true } as never)).toThrow('opts')
})
