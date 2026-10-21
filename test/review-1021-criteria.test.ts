/** Reproductions from the 21.10 review of the criteria functions and wildcard patterns. */
import { describe, expect, it } from 'vitest'
import { compute, createSession, type BookInput } from '../src/index.js'
import { findMatcher, findReplacer } from '../src/text.js'
import './locales.js'

type Cells = BookInput['sheets'][number]['cells']
const book = (cells: Cells, locale?: string): BookInput => ({ ...(locale && { locale }), sheets: [{ name: 'S', cells }] })

/** Milliseconds a call takes. */
function took(f: () => unknown): number {
  const t0 = performance.now()
  f()
  return performance.now() - t0
}

describe.each(['closures', 'eval'] as const)('review 21.10 criteria, %s', (execution) => {
  it.each(['en-US', 'ru-RU'])('an array in a range of criteria under an operator is an array of #VALUE!, %s', (locale) => {
    const VALUE = { error: 'VALUE' }
    const cases: [formula: string, spill: string | undefined, value: unknown][] = [
      ['COUNTIF({1},">0")*1', undefined, VALUE],
      ['-SUMIF({1},">0")', undefined, VALUE],
      ['COUNTIF({1,2},">0")*1', 'A1:B1', VALUE],
      ['COUNTIF({1,2},">0")&""', 'A1:B1', VALUE],
      ['IF(COUNTIF({1,2},">0"),1,2)', 'A1:B1', VALUE],
      ['COUNTIF({1,2},">0")=1', 'A1:B1', VALUE],
      ['COUNTIF(B5:B6*1,">0")*1', 'A1:A2', VALUE],
      ['SUMIFS(B5:B6,B5:B6*1,">0")*1', 'A1:A2', VALUE],
      // the range summed too
      ['SUMIFS(B5:B6*1,B5:B6,">0")*1', 'A1:A2', VALUE],
      ['SUMIF(B5:B6,">0",B5:B6*1)*1', 'A1:A2', VALUE],
      ['MAXIFS({1;2},B5:B6,">0")*1', 'A1:A2', VALUE],
      ['LET(x,{1,2},COUNTIF(x,">0")+0)', 'A1:B1', VALUE],
      ['IFERROR(COUNTIF({1},">0")*1,7)', undefined, 7],
      // a reference there stays a range of criteria
      ['SUMIF(IF(TRUE,B5:B6),">0")*1', undefined, 8],
      ['SUMIF(INDEX(B5:B6,0,1),">0")*1', undefined, 8],
    ]
    for (const [formula, spill, value] of cases) {
      const b = book({ B5: { value: 3 }, B6: { value: 5 }, A1: { formula } }, locale)
      const cell = compute(b, { execution, display: false }).sheets[0].cells.A1
      expect([cell.value, cell.spill], formula).toEqual([value, spill])
    }
    const s = createSession(book({ B5: { value: 3 }, A1: { formula: 'COUNTIF({1},">0")*1' } }, locale), { execution })
    expect(s.setCells([{ sheet: 'S', cell: 'A2', input: { formula: 'COUNTIF({1,2},">0")+0' } }]).changes.find((c) => c.cell === 'A2')?.value).toEqual(VALUE)
  })

  it('wildcards cost a search per part, not a backtracking over the text per *', () => {
    const long = 'a'.repeat(4000)
    const cells: Cells = { A1: { value: long }, A2: { value: long + 'c' }, A3: { value: 'a'.repeat(30000) } }
    const formulas: [string, unknown][] = [
      // a regex with [\s\S]* per * took about 8 s here
      ['COUNTIF(A1:A2,"*a*a*c")', 1],
      ['COUNTIF(A1:A3,"<>*a*a*a*c")', 2],
      ['SUMPRODUCT(--ISNUMBER(MATCH("*a*a*c",A1:A3,0)))', 1],
      ['XMATCH("*a*a*c",A1:A3,2)', 2],
      ['IFERROR(VLOOKUP("a*a*a*b",A1:A3,1,FALSE),"none")', 'none'],
      // SEARCH over 500 characters took about 16 s
      ['IFERROR(SEARCH("*a*a*c",A3),0)', 0],
      ['SEARCH("a*a*c",A2)', 1],
      ['SEARCH("?a*c",A2,3990)', 3990],
    ]
    for (const [i, [formula]] of formulas.entries()) cells[`B${i + 1}`] = { formula }
    let out: Cells = {}
    expect(took(() => (out = compute(book(cells), { execution, display: false }).sheets[0].cells))).toBeLessThan(1000)
    expect(formulas.map((_, i) => out[`B${i + 1}`].value)).toEqual(formulas.map(([, v]) => v))
  })

  it('Find and Replace with wildcards over a long text cost a search per part', () => {
    const text = 'a'.repeat(30000)
    const s = createSession(book({ A1: { value: text }, A2: { value: 'xayazc' } }), { execution })
    const found = (query: string) => s.findCells([{ sheet: 'S', range: 'A1:A2' }], query, { wildcards: true, fields: [] }).cells.map((c) => c.cell)
    let cells: string[] = []
    expect(took(() => (cells = found('*a*a*c')))).toBeLessThan(1000)
    expect([cells, found('a*z?'), found('a?a*')]).toEqual([['A2'], ['A2'], ['A1', 'A2']])
    const find = findMatcher('*a*a*c', { wildcards: true })!
    const replace = findReplacer('a*a*c', '#', { wildcards: true })!
    expect(took(() => [find(text), replace(text)])).toBeLessThan(1000)
    expect([find(text), replace(text), replace('xaabcac'), findReplacer('a*', '#', { wildcards: true })!('baab')]).toEqual([false, null, 'x#ac', 'b#'])
  })
})
