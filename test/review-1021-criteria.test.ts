/** Reproductions from the 21.10 review of the criteria functions and wildcard patterns. */
import { describe, expect, it } from 'vitest'
import { compute, createSession, type BookInput } from '../src/index.js'
import './locales.js'

type Cells = BookInput['sheets'][number]['cells']
const book = (cells: Cells, locale?: string): BookInput => ({ ...(locale && { locale }), sheets: [{ name: 'S', cells }] })

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
})
