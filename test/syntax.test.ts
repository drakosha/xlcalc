import { describe, expect, it } from 'vitest'
import { canonical, parse, validate, validateName } from '../src/syntax'

const ctx = { sheets: ['Данные', 'S'], names: ['data_amt', 'one'] }
const isRange = (n: string) => (n === 'data_amt' ? true : n === 'one' ? false : undefined)
const code = (f: string, c = ctx) => {
  const r = validate(f, c, isRange)
  return r.ok ? 'ok' : r.code
}

describe('разбор формул', () => {
  it('канон: лист в кавычках, функции заглавными, удвоение кавычек, без пробелов', () => {
    expect(canonical(parse("sumifs(Данные!D:D, 'Данные'!B:B, \"он \"\"да\"\"\")"))).toBe(`SUMIFS('Данные'!D:D,'Данные'!B:B,"он ""да""")`)
    expect(canonical(parse("'It''s'!$A$1+1"))).toBe("'It''s'!$A$1+1")
    expect(canonical(parse('-2^2+5%'))).toBe('-2^2+5%')
    expect(canonical(parse('IF(true,1,2)'))).toBe('IF(TRUE,1,2)')
  })

  it('ссылки трёх видов и имя', () => {
    const a = parse("SUM('Данные'!$D$2:$D$3401,Данные!D:D,data_amt,1:1)")
    expect(a.k === 'fn' && a.args.map((x) => (x.k === 'ref' ? x.ref.kind : x.k))).toEqual(['range', 'cols', 'name', 'rows'])
  })

  it('коды проверки', () => {
    expect(code('')).toBe('empty')
    expect(code('SUM(A1')).toBe('syntax')
    // форма ошибки из Python-книги пробы: запятая вне функции
    expect(code('SUM(A1),SUM(B1)')).toBe('syntax')
    expect(code('1+' + '1+'.repeat(5000) + '1')).toBe('too_long')
    expect(code('('.repeat(70) + '1' + ')'.repeat(70))).toBe('too_deep')
    expect(code('[1]Лист!A1')).toBe('external_ref')
    expect(code('VLOOKUP(1,A1:B2,2)')).toBe('unknown_function')
    expect(code('SUM(Нет!A1)')).toBe('unknown_sheet')
    expect(code('SUM(nope)')).toBe('unknown_name')
    expect(code('SUM(данные!A1)')).toBe('ok') // лист без учёта регистра
  })

  it('число аргументов', () => {
    expect(code('ROUND(1)')).toBe('syntax')
    expect(code('SUMIFS(A:A,B:B)')).toBe('syntax')
    expect(code('COUNTIFS(A:A,1,B:B)')).toBe('syntax')
    expect(code('IF(1,2,3,4)')).toBe('syntax')
    const r = validate('ROUND(1)', ctx)
    expect(!r.ok && r.message).toBe('ROUND takes 2 arguments, got 1')
  })

  it('диапазон в скалярном месте: unsupported', () => {
    expect(code('A1:A3*2')).toBe('unsupported')
    expect(code('ROUND(A1:A3,0)')).toBe('unsupported')
    expect(code('A1:A3')).toBe('unsupported')
    expect(code('SUMIFS(A:A,B:B,C1:C2)')).toBe('unsupported')
    expect(code('data_amt+1')).toBe('unsupported')
    expect(code('one+1')).toBe('ok')
    expect(code('A1:A1*2')).toBe('ok')
    expect(code('SUM(A1:A3)*2')).toBe('ok')
    expect(code('SUMIFS(A:A,B:B,C1)')).toBe('ok')
    expect(code('XLOOKUP(A1,B:B,C:C)')).toBe('ok')
  })

  it('ошибка указывает место', () => {
    const r = validate('SUM(A1,,', ctx)
    expect(r.ok).toBe(false)
    if (!r.ok) expect(r.at).toBeGreaterThan(0)
  })

  it('результат: канон, функции, листы, имена', () => {
    const r = validate("sum(данные!a1, data_amt)", ctx, isRange)
    expect(r.ok && { c: r.canonical, f: r.functions, s: r.sheets, n: r.names }).toEqual({ c: "SUM('данные'!A1,data_amt)", f: ['SUM'], s: ['данные'], n: ['data_amt'] })
  })

  it('внешние функции: только литералы, вызовы списком', () => {
    const c = { ...ctx, externalFunctions: ['BALANCE'] }
    const r = validate('BALANCE("2025-01")+balance("2025-02",-1)', c)
    expect(r.ok && r.externals).toEqual(['BALANCE("2025-01")', 'BALANCE("2025-02",-1)'])
    expect(code('BALANCE(A1)', c)).toBe('unsupported')
    expect(code('BALANCE("x")')).toBe('unknown_function')
  })

  it('имена: латиница, не адрес, не R1C1, не TRUE', () => {
    expect(validateName('data_amt').ok).toBe(true)
    for (const bad of ['Q1', 'R1C1', 'TRUE', 'Сумма', '1x', 'a b']) expect(validateName(bad).ok).toBe(false)
  })
})
