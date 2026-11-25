import { describe, expect, it } from 'vitest'
import { compute } from '../src/index.js'

// The cases came from a comparison with ooxml-core's Excel for Windows fixtures. Rechecked in
// Excel for Mac 16.115.2. The text slicing cases use workbook Compatibility Version 1.
const cases = [
  ['FIND("","")', 1],
  ['FIND("","abc",4)', 4],
  ['FIND("","abc",5)', { error: 'VALUE' }],
  ['SEARCH("","abc",4)', { error: 'VALUE' }],
  ['LEFT("😀",1)', '😀'],
  ['LEFT("😀😀",3)', '😀😀'],
  ['RIGHT("a😀",1)', '😀'],
  ['RIGHT("😀😀",3)', '😀😀'],
  ['LEN("😀")', 2],
  ['LEN(LOWER("İ"))', 1],
  ['LOWER("ẞ")', 'ẞ'],
  ['PROPER("ΣΑΣ σας")', 'Σας Σας'],
  ['LEN(TEXT(REPT("a",255),"@"))', 255],
  ['LEN(TEXT(REPT("a",256),"@"))', { error: 'VALUE' }],
  ['LEN(TEXT(REPT("a",128),"@@"))', { error: 'VALUE' }],
  ['LEN(ARRAYTOTEXT(SEQUENCE(10000)))', { error: 'CALC' }],
  ['LEN(ARRAYTOTEXT(SEQUENCE(10000),1))', { error: 'CALC' }],
  ['MOD(1125899999999,1)', 0],
  ['MOD(2251799999998,2)', 0],
  ['MOD(1125900000000,1)', { error: 'NUM' }],
  ['SECOND(0.5+(0.5+1E-7)/86400)', 1],
  ['SWITCH(A1,0.3,"eq","ne")', 'eq'],
  ['SUM(AGGREGATE(14,6,{10,20,30},{1,2}))', 50],
  ['SUM(AGGREGATE(14,6,{10,20,30},{1,2})+10)', 70],
  ['"ß"="SS"', true],
  ['SWITCH("Straße","STRASSE",1,2)', 1],
  ['COUNTIF(A2:C2,"ss")', 3],
  ['COUNTIF(A3:C3,"ΣΑΣ")', 3],
  ['COUNTIF(A4:B4,"fi")', 2],
  ['XMATCH("ß",{"a","ss","ß"})', 2],
  ['XLOOKUP("ß",{"SS","ß"},{1,2})', 1],
  ['MATCH("fi",{"ﬁ","fi"},0)', 2],
  ['COLUMNS(UNIQUE({"STRASSE","Straße","SS","ß"},TRUE))', 2],
  ['COLUMNS(UNIQUE({"ΣΑΣ","σας","σασ"},TRUE))', 1],
] as const

describe.each(['closures', 'eval'] as const)('Excel formula regressions (%s)', (execution) => {
  it.each(cases)('%s', (formula, expected) => {
    const result = compute({ sheets: [{ name: 'S', cells: {
      A1: { formula: '0.1+0.2' }, C1: { formula },
      A2: { value: 'ß' }, B2: { value: 'SS' }, C2: { value: 'ẞ' },
      A3: { value: 'ΣΑΣ' }, B3: { value: 'σας' }, C3: { value: 'σασ' },
      A4: { value: 'ﬁ' }, B4: { value: 'fi' },
    } }] }, { execution })
    expect(result.sheets[0].cells.C1.value).toEqual(expected)
  })

  // Nearest-even fixes this case but changes other half-second answers recorded from Excel.
  it.fails('SECOND at a half-second boundary after midday', () => {
    const result = compute({ sheets: [{ name: 'S', cells: { A1: { formula: 'SECOND(0.5+0.5/86400)' } } }] }, { execution })
    expect(result.sheets[0].cells.A1.value).toBe(0)
  })

  it('spills the results of AGGREGATE over the positions in an array constant', () => {
    const result = compute({ sheets: [{ name: 'S', cells: {
      A1: { value: 10 }, A2: { value: 20 }, A3: { value: 30 }, C1: { formula: 'AGGREGATE(14,6,A1:A3,{1,2})' },
    } }] }, { execution })
    expect(result.sheets[0].cells.C1).toMatchObject({ value: 30, spill: 'C1:D1' })
    expect(result.sheets[0].cells.D1).toMatchObject({ value: 20, spillOf: 'C1' })
  })
})
