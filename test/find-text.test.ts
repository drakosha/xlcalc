/** Find and Replace over a host's texts (`findMatcher`, `findReplacer` of `xlcalc/text`) answer as a session's `findCells` and `replaceInCells` over the same texts. */
import { expect, it } from 'vitest'
import { compute, createSession, type CellValue } from '../src/index.js'
import { compareValues, findMatcher, findReplacer, valueComparator } from '../src/text.js'
import { ruRU } from '../src/locales/ru-RU.js'
import { registerLocale } from '../src/index.js'

registerLocale(ruRU)

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

it('compareValues sorts as the Sort command does: numbers, text, booleans, errors equal, blanks last either way; SORT orders errors by ERROR.TYPE', () => {
  const vals: CellValue[] = [null, 'b', true, { error: 'NA' }, 2, 'A', false, -1, { error: 'DIV0' }]
  const sorted = (descending: boolean) => [...vals].sort((a, b) => compareValues(a, b, { descending }))
  expect(sorted(false)).toEqual([-1, 2, 'A', 'b', false, true, { error: 'NA' }, { error: 'DIV0' }, null])
  expect(sorted(true)).toEqual([{ error: 'NA' }, { error: 'DIV0' }, true, false, 'b', 'A', 2, -1, null])
  const input = (v: CellValue) => (v === null ? {} : typeof v === 'object' ? { formula: v.error === 'NA' ? 'NA()' : '1/0' } : { value: v })
  const cells = { ...Object.fromEntries(vals.map((v, i) => [`A${i + 1}`, input(v)])), C1: { formula: 'SORT(A1:A9)' } }
  const out = compute({ sheets: [{ name: 'S', cells }] }).sheets[0].cells
  // SORT gives a blank as 0, and #DIV/0! before #N/A
  const bySort = sorted(false).map((v) => v ?? 0)
  bySort.splice(6, 2, { error: 'DIV0' }, { error: 'NA' })
  expect(Array.from({ length: 9 }, (_, i) => out[`C${i + 1}`].value)).toEqual(bySort)
})

it('valueComparator answers as compareValues on every pair of mixed values, with descending and without, in a locale\'s order of text', () => {
  const vals: CellValue[] = [null, 0, -1.5, 2, 1e300, '', 'a', 'A', 'b', 'Б', 'ё', 'Е', '10', true, false, { error: 'NA' }, { error: 'DIV0' }, { error: 'VALUE' }]
  for (const opts of [undefined, {}, { descending: false }, { descending: true }]) {
    const cmp = valueComparator(opts)
    for (const a of vals) for (const b of vals) expect([a, b, opts, cmp(a, b)]).toEqual([a, b, opts, compareValues(a, b, opts)])
  }
  // a locale's order of text: ru-RU's is en-US's, Latin first, as Excel for Windows sorts; one that puts Cyrillic first
  const cyrillicFirst = { ...ruRU, tag: 'ru-XC', regional: { ...ruRU.regional!, collation: 'ru' } }
  for (const opts of [{ locale: 'ru-RU' }, { locale: ruRU, descending: true }, { locale: cyrillicFirst }]) {
    const cmp = valueComparator(opts)
    for (const a of vals) for (const b of vals) expect([a, b, cmp(a, b)]).toEqual([a, b, compareValues(a, b, opts)])
  }
  const sorted = (opts: object) => ['b', 'а', 'A', 'Б', 'ё', 'е'].sort(valueComparator(opts))
  expect(sorted({ locale: ruRU })).toEqual(sorted({}))
  expect(sorted({})).toEqual(['A', 'b', 'а', 'Б', 'е', 'ё'])
  expect(sorted({ locale: cyrillicFirst })).toEqual(['а', 'Б', 'е', 'ё', 'A', 'b'])
  expect(() => valueComparator({ locale: 'xx-YY' })).toThrow('opts.locale')
  expect(() => valueComparator({ desc: true } as never)).toThrow('opts')
  expect(() => valueComparator({ descending: 1 } as never)).toThrow('opts.descending')
})
