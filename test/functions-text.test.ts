/**
 * Text functions beyond the Excel case set (`bench/coverage/sets/text.ts`): the 32,767-character limit,
 * several delimiters in TEXTBEFORE and TEXTAFTER, VALUETOTEXT over a range, astral characters.
 */
import { describe, expect, it } from 'vitest'
import { validateFormula } from '../src/index.js'
import { err, valuesAt } from './formula-values.js'

const values = valuesAt('2026-10-06T12:00:00Z')

describe('text functions beyond the Excel cases', () => {
  it('& and CONCATENATE cut the text at 32,767 characters, other functions refuse a longer result', () => {
    expect(values(['LEN(REPT("a",20000)&REPT("b",20000))', 'RIGHT(CONCATENATE(REPT("a",20000),"xyz",REPT("b",20000)),1)', 'LEN(REPT("a",32767)&"")'])).toEqual([32767, 'b', 32767])
  })

  it('results longer than 32,767 characters are #CALC! from CONCAT and TEXTJOIN, #VALUE! from other functions', () => {
    expect(values(['CONCAT(REPT("a",20000),REPT("b",20000))', 'TEXTJOIN(",",TRUE,REPT("a",20000),REPT("b",20000))', 'SUBSTITUTE(REPT("a",20000),"a","bb")', 'REPLACE(REPT("a",30000),1,0,REPT("b",3000))', 'LEN(TEXTJOIN(",",TRUE,REPT("a",16383),REPT("b",16383)))'])).toEqual([
      err('CALC'),
      err('CALC'),
      err('VALUE'),
      err('VALUE'),
      32767,
    ])
  })

  it('TEXTBEFORE and TEXTAFTER take several delimiters: the nearest one counts', () => {
    expect(validateFormula('TEXTBEFORE("a-b",A1:A2)').ok).toBe(true)
    expect(values(['TEXTBEFORE("a-b_c",{"_","-"})', 'TEXTAFTER("a-b_c",{"_","-"},2)', 'TEXTBEFORE("a-b_c",A1:A2,-1)', 'TEXTAFTER("a-b",{"x","y"})'], { A1: { value: '_' }, A2: { value: '-' } })).toEqual(['a', 'c', 'a-b', err('NA')])
  })

  it('VALUETOTEXT over a range of several cells is lifted: an array, #SPILL! in a cell', () => {
    expect(values(['VALUETOTEXT(A1:A2)', 'VALUETOTEXT(A1)', 'CONCAT(VALUETOTEXT(A1:A2,1))'], { A1: { value: 1.5 }, A2: { value: 'x' } })).toEqual([err('SPILL'), '1.5', '1.5"x"'])
  })

  it('PROPER, UPPER and MID keep astral characters whole (UTF-16 units count, as in Excel)', () => {
    expect(values(['PROPER("😀abc def")', 'UPPER("😀ß")', 'LEN("😀")', 'MID("a😀b",2,2)'])).toEqual(['😀Abc Def', '😀ß', 2, '😀'])
  })

  it('SEARCH from a start inside a surrogate pair answers a position at the start or after it, the same pattern again too', () => {
    const calls = ['SEARCH("?","😀a",2)', 'SEARCH("?","a😀b",3)', 'SEARCH("*","😀",2)', 'SEARCH("?a","x😀a",3)']
    expect(values([...calls, ...calls])).toEqual([2, 3, 2, 3, 2, 3, 2, 3])
  })
})
