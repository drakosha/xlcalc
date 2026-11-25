import { describe, expect, it } from 'vitest'
import { compute } from '../src/index.js'

// Excel cases from the ooxml-core comparison and its follow-up probes on Mac and Windows.
// The text slicing cases use workbook Compatibility Version 1.
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
  ['SECOND(0.5+0.5/86400)', 0],
  ['SECOND(0.5+(0.5+1E-7)/86400)', 1],
  ['SECOND(0.5+9.5/86400)', 9],
  ['SECOND(10.5/86400)', 11],
  ['BINOM.INV(3,0.25,0.421875)', 0],
  ['BINOM.INV(3,0.25,0.421875000000001)', 1],
  ['BINOM.INV(2,0.25,0.5625)', 1],
  ['BINOM.INV(10,0.5,0.5^10)', 1],
  ['BINOM.INV(1,0.67,A5)', 1],
  ['BINOM.INV(1,0.67,B5)', 0],
  ['BINOM.INV(512,A6,B6)', 1],
  ['CRITBINOM(3,0.25,0.421875)', 0],
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
      A5: { value: 0.3299999999999999 }, B5: { value: 0.32999999999999985 },
      A6: { value: 0.7073477915737445 }, B6: { value: 5.917313071032519e-274 },
    } }] }, { execution })
    expect(result.sheets[0].cells.C1.value).toEqual(expected)
  })

  it('spills the results of AGGREGATE over the positions in an array constant', () => {
    const result = compute({ sheets: [{ name: 'S', cells: {
      A1: { value: 10 }, A2: { value: 20 }, A3: { value: 30 }, C1: { formula: 'AGGREGATE(14,6,A1:A3,{1,2})' },
    } }] }, { execution })
    expect(result.sheets[0].cells.C1).toMatchObject({ value: 30, spill: 'C1:D1' })
    expect(result.sheets[0].cells.D1).toMatchObject({ value: 20, spillOf: 'C1' })
  })
})
