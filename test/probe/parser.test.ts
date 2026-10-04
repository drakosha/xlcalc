import { describe, expect, it } from 'vitest'
import { WHITELIST } from '../../src/probe/functions'
import { canonical, parse, validateFormula } from '../../src/probe/parser'
import { validateName } from '../../src/probe/prep'

const ctx = { sheets: ['Данные', 'S'], names: ['data_amt'], functions: WHITELIST }
const code = (f: string) => {
  const r = validateFormula(f, ctx)
  return r.ok ? 'ok' : r.code
}

describe('наш разбор формул', () => {
  it('канон: лист в кавычках, функции заглавными, удвоение кавычек, без пробелов', () => {
    expect(canonical(parse("sumifs(Данные!D:D, 'Данные'!B:B, \"он \"\"да\"\"\")"))).toBe(`SUMIFS('Данные'!D:D,'Данные'!B:B,"он ""да""")`)
    expect(canonical(parse("'It''s'!$A$1+1"))).toBe("'It''s'!$A$1+1")
    expect(canonical(parse('-2^2+5%'))).toBe('-2^2+5%')
    expect(canonical(parse('IF(true,1,2)'))).toBe('IF(TRUE,1,2)')
  })

  it('разбор: ссылки трёх видов и имя', () => {
    const a = parse("SUM('Данные'!$D$2:$D$3401,Данные!D:D,data_amt,1:1)")
    expect(a.k === 'fn' && a.args.map((x) => (x.k === 'ref' ? x.ref.kind : x.k))).toEqual(['range', 'cols', 'name', 'rows'])
  })

  it('коды проверки', () => {
    expect(code('')).toBe('empty')
    expect(code('SUM(A1')).toBe('syntax')
    // форма ошибки из Python-книги пробы: запятая вне функции
    expect(code('A1,B1')).toBe('syntax')
    expect(code('1+' + '1+'.repeat(5000) + '1')).toBe('too_long')
    expect(code('('.repeat(70) + '1' + ')'.repeat(70))).toBe('too_deep')
    expect(code('[1]Лист!A1')).toBe('external_ref')
    expect(code('VLOOKUP(1,A1:B2,2)')).toBe('unknown_function')
    expect(code('SUM(Нет!A1)')).toBe('unknown_sheet')
    expect(code('SUM(nope)')).toBe('unknown_name')
    expect(code('SUM(данные!A1)')).toBe('ok') // лист без учёта регистра
  })

  it('ошибка указывает место', () => {
    const r = validateFormula('SUM(A1,,', ctx)
    expect(r.ok).toBe(false)
    if (!r.ok) expect(r.at).toBeGreaterThan(0)
  })

  it('имена: латиница, не адрес, не R1C1, не TRUE', () => {
    expect(validateName('data_amt').ok).toBe(true)
    for (const bad of ['Q1', 'R1C1', 'TRUE', 'Сумма', '1x']) expect(validateName(bad).ok).toBe(false)
  })
})
