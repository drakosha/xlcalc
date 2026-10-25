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
    let out: Record<string, { value: unknown }> = {}
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

describe.each(['closures', 'eval'] as const)('SUMIF and AVERAGEIF with a sum range resized past the sheet edge, %s', (execution) => {
  // Excel for Mac and for Windows, the set low-values-held: the cells that exist, positions past the edge blank
  it('add or average the cells that exist', () => {
    const cases: [string, unknown][] = [
      ['SUMIF(A1:A3,">0",B1048575)', 30],
      ['SUMIF(A1:A3,">2",B1048575)', 0],
      ['SUMIF(A1:A3,">1",B1048575)', 20],
      ['SUMIF(A1:A3,">0",B1048576)', 20],
      ['SUMIF(H1:J1,">0",XFC1)', 3000],
      ['AVERAGEIF(A1:A3,">0",B1048575)', 15],
      ['AVERAGEIF(A1:A3,">2",B1048575)', { error: 'DIV0' }],
    ]
    const cells: Cells = { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 }, B1048575: { value: 10 }, B1048576: { value: 20 }, H1: { value: 1 }, I1: { value: 2 }, J1: { value: 3 }, XFC1: { value: 1000 }, XFD1: { value: 2000 } }
    cases.forEach(([f], i) => (cells[`D${i + 1}`] = { formula: f }))
    const r = compute(book(cells), { execution }).sheets[0].cells
    expect(cases.map((_, i) => r[`D${i + 1}`].value)).toEqual(cases.map(([, v]) => v))
  })

  it('follow an edit of a cell at the edge in a session', () => {
    const s = createSession(book({ A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 }, B1048576: { value: 20 }, D1: { formula: 'SUMIF(A1:A3,">0",B1048575)' } }), { execution })
    s.setCells([{ sheet: 'S', cell: 'B1048576', input: { value: 5 } }])
    expect(s.getCells('S', 'D1').D1.value).toBe(5)
    s.close()
  })
})

describe.each(['closures', 'eval'] as const)('parsed criteria kept by value, %s', (execution) => {
  it('keep a number, its text, a boolean and its text apart', () => {
    const cells: Cells = { A1: { value: 1 }, A2: { value: '1' }, A3: { value: true }, A4: { value: 'TRUE' }, A5: { value: 0 }, B1: { value: 1 }, B2: { value: '1' }, B3: { value: true }, B4: { value: 'TRUE' }, B5: { formula: '-0' } }
    for (let r = 1; r <= 5; r++) cells[`C${r}`] = { formula: `COUNTIF(A1:A5,B${r})` }
    const out = compute(book(cells), { execution }).sheets[0].cells
    expect([1, 2, 3, 4, 5].map((r) => out[`C${r}`].value)).toEqual([2, 2, 1, 1, 1])
  })
})
