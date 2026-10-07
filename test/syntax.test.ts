import { describe, expect, it } from 'vitest'
import { canonical, parse, validate, validateName } from '../src/syntax.js'

const ctx = { sheets: ['Данные', 'S'], names: ['data_amt', 'one'] }
const code = (f: string, c = ctx) => {
  const r = validate(f, c)
  return r.ok ? 'ok' : r.code
}

describe('formula parsing', () => {
  it('canonical form: quoted sheet, upper-case functions, doubled quotes, no whitespace', () => {
    expect(canonical(parse("sumifs(Данные!D:D, 'Данные'!B:B, \"он \"\"да\"\"\")"))).toBe(`SUMIFS('Данные'!D:D,'Данные'!B:B,"он ""да""")`)
    expect(canonical(parse("'It''s'!$A$1+1"))).toBe("'It''s'!$A$1+1")
    expect(canonical(parse('-2^2+5%'))).toBe('-2^2+5%')
    expect(canonical(parse('IF(true,1,2)'))).toBe('IF(TRUE,1,2)')
  })

  it('number literals keep 15 significant digits, the rest is dropped (as Excel reads them)', () => {
    const num = (f: string) => {
      const a = parse(f)
      return a.k === 'num' ? a.v : NaN
    }
    expect([num('123456789012345678'), num('1.99999999999999999'), num('.5'), num('5.'), num('1234567890123456789E-5'), num('0.000000000000000012345678901234567')]).toEqual([
      123456789012345000, 1.99999999999999, 0.5, 5, 12345678901234.5, 1.23456789012345e-17,
    ])
  })

  it('three kinds of references and a name', () => {
    const a = parse("SUM('Данные'!$D$2:$D$3401,Данные!D:D,data_amt,1:1)")
    expect(a.k === 'fn' && a.args.map((x) => (x.k === 'ref' ? x.ref.kind : x.k))).toEqual(['range', 'cols', 'name', 'rows'])
  })

  it('check codes', () => {
    expect(code('')).toBe('empty')
    expect(code('SUM(A1')).toBe('syntax')
    // a comma outside a function call
    expect(code('SUM(A1),SUM(B1)')).toBe('syntax')
    expect(code('1+' + '1+'.repeat(5000) + '1')).toBe('too_long')
    expect(code('('.repeat(70) + '1' + ')'.repeat(70))).toBe('too_deep')
    expect(code('[1]Лист!A1')).toBe('external_ref')
    expect(code('INDIRECT("A1")')).toBe('unknown_function')
    expect(code('SUM(Нет!A1)')).toBe('unknown_sheet')
    expect(code('SUM(nope)')).toBe('unknown_name')
    expect(code('SUM(данные!A1)')).toBe('ok') // sheet names are case-insensitive
  })

  it('argument counts', () => {
    expect(code('ROUND(1)')).toBe('syntax')
    expect(code('SUMIFS(A:A,B:B)')).toBe('syntax')
    expect(code('COUNTIFS(A:A,1,B:B)')).toBe('syntax')
    expect(code('IF(1,2,3,4)')).toBe('syntax')
    // Excel does not load these (measured: the core set's arity cases)
    expect(code('SORTBY(A1:A3,B1:B3,-1,C1:C3)')).toBe('syntax')
    expect(code('SORTBY(A1:A3,B1:B3,-1,C1:C3,1,A1:A3)')).toBe('syntax')
    expect(code('SORTBY(A1:A3,B1:B3,-1,C1:C3,1)')).toBe('ok')
    expect(code(`SWITCH(1,${'1,1,'.repeat(127).slice(0, -1)})`)).toBe('syntax')
    expect(code(`SWITCH(1,${'1,1,'.repeat(126)}0)`)).toBe('ok')
    expect(validate('SORTBY(A1:A3,B1:B3,-1,C1:C3)', ctx)).toMatchObject({ message: 'SORTBY takes 2, or odd 3 to 253 arguments, got 4' })
    const r = validate('ROUND(1)', ctx)
    expect(!r.ok && r.message).toBe('ROUND takes 2 arguments, got 1')
  })

  it('a result that is an array is valid: it spills', () => {
    for (const f of ['A1:A3*2', 'ROUND(A1:A3,0)', 'A1:A3', 'SUMIFS(A:A,B:B,C1:C2)', 'data_amt+1', '{1,2}', 'ROW(A1:A3)', 'IF({TRUE,FALSE},1,2)', 'one+1', 'SUM(A1:A3*2)', '{1}', 'INDEX(A1:C3,0,1)'])
      expect(code(f), f).toBe('ok')
  })

  it('array constants: rows by ;, values by , ; numbers with a sign, text, booleans, errors', () => {
    expect(canonical(parse('SUM({1,-2.5;"a""b",true;#N/A,+3})'))).toBe('SUM({1,-2.5;"a""b",TRUE;#N/A,3})')
    for (const f of ['{1,2;3}', '{}', '{A1}', '{1+2}', '{{1}}', '{1;}', 'SUM({1,2)', '1;2']) expect(code(f), f).toBe('syntax')
    const r = validate('{1,2;3}', ctx)
    expect(!r.ok && r.message).toBe('the rows of an array constant must have the same number of values')
  })

  it('an error points at the position', () => {
    const r = validate('SUM(A1,,', ctx)
    expect(r.ok).toBe(false)
    if (!r.ok) expect(r.at).toBeGreaterThan(0)
  })

  it('result: canonical text, functions, sheets, names', () => {
    const r = validate("sum(данные!a1, data_amt)", ctx)
    expect(r.ok && { c: r.canonical, f: r.functions, s: r.sheets, n: r.names }).toEqual({ c: "SUM('данные'!A1,data_amt)", f: ['SUM'], s: ['данные'], n: ['data_amt'] })
  })

  it('external functions: literals only, calls listed', () => {
    const c = { ...ctx, externalFunctions: ['BALANCE'] }
    const r = validate('BALANCE("2025-01")+balance("2025-02",-1)', c)
    expect(r.ok && r.externals).toEqual(['BALANCE("2025-01")', 'BALANCE("2025-02",-1)'])
    expect(code('BALANCE(A1)', c)).toBe('unsupported')
    expect(code('BALANCE("x")')).toBe('unknown_function')
  })

  it('names: Latin, not an address, not R1C1, not TRUE', () => {
    expect(validateName('data_amt').ok).toBe(true)
    for (const bad of ['Q1', 'R1C1', 'TRUE', 'Сумма', '1x', 'a b']) expect(validateName(bad).ok).toBe(false)
  })
})
