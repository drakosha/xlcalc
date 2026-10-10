/**
 * Locales of formula text (`src/locale/`): translation both ways, what stays as it is, errors with positions
 * in the given text, checking typed text, locale objects of one's own and malformed ones. Excel's own text
 * is the reference in `locale-excel.test.ts`, the round trip over all case formulas in `locale-roundtrip.test.ts`.
 */
import { describe, expect, it } from 'vitest'
import { enUS, InputError, ruRU, toInvariant, toLocal, validateFormula, validateLocalFormula, type Locale } from '../src/index.js'

const local = (f: string, loc: Locale = ruRU) => {
  const r = toLocal(f, loc)
  return r.ok ? r.formula : `${r.code} at ${r.at}: ${r.message}`
}
const invariant = (t: string, loc: Locale = ruRU) => {
  const r = toInvariant(t, loc)
  return r.ok ? r.formula : `${r.code} at ${r.at}: ${r.message}`
}

describe('translation between the invariant syntax and ru-RU', () => {
  it('function names, separators and decimal commas both ways', () => {
    expect(invariant('=СУММ(A1:A3;1,5)')).toBe('SUM(A1:A3,1.5)')
    expect(local('SUM(A1:A3,1.5)')).toBe('СУММ(A1:A3;1,5)')
    expect(local('ROUND(XLOOKUP(1.25,A:A,B:B),-1.5E-07)')).toBe('ОКРУГЛ(ПРОСМОТРX(1,25;A:A;B:B);-1,5E-07)')
  })

  it('TRUE and FALSE, error literals and array constants', () => {
    expect(local('IF(TRUE,{1.5,2;3,4},#N/A)&FALSE()')).toBe('ЕСЛИ(ИСТИНА;{1,5;2:3;4};#Н/Д)&ЛОЖЬ()')
    expect(invariant('ЕСЛИ(ИСТИНА;{1,5;2:3;4};#Н/Д)&ЛОЖЬ()')).toBe('IF(TRUE,{1.5,2;3,4},#N/A)&FALSE()')
    expect(local('ISERROR(#DIV/0!)+ISNA(#VALUE!)+{"a",TRUE;#REF!,-1}')).toBe('ЕОШИБКА(#ДЕЛ/0!)+ЕНД(#ЗНАЧ!)+{"a";ИСТИНА:#ССЫЛКА!;-1}')
  })

  it('strings, sheet names, defined names, references and whitespace stay as they are', () => {
    expect(local(`SUMIF(items,"a;b,c 1.5",'Лист, 1.5'!B:B) + Данные!$A$1:A2 * tax_rate`)).toBe(`СУММЕСЛИ(items;"a;b,c 1.5";'Лист, 1.5'!B:B) + Данные!$A$1:A2 * tax_rate`)
    expect(local('SUM( 1.5 , 2 )')).toBe('СУММ( 1,5 ; 2 )')
    expect(local("'It''s'!A1&\"say \"\"1,5\"\"\"")).toBe("'It''s'!A1&\"say \"\"1,5\"\"\"")
  })

  it('the union operator is the list separator, the intersection a space', () => {
    expect(local('SUM((A1,B2),C3 C4)')).toBe('СУММ((A1;B2);C3 C4)')
  })

  it('typed text is read in any case; English names, TRUE/FALSE and error texts are read too', () => {
    expect(invariant('=сумм(1;истина;#н/д)')).toBe('SUM(1,TRUE,#N/A)')
    expect(invariant('=SUM(1;TRUE;#N/A)')).toBe('SUM(1,TRUE,#N/A)')
    expect(invariant('{true;ложь}')).toBe('{TRUE,FALSE}')
  })

  it('functions the locale does not know pass through: host functions, prefixed names', () => {
    expect(invariant('BALANCE("2026-01";1,5)*2')).toBe('BALANCE("2026-01",1.5)*2')
    expect(local('_xlfn.ENCODEURL("a b")')).toBe('_xlfn.ENCODEURL("a b")')
  })

  it("a table's column named with a list separator keeps it, in brackets after @ where it is the locale's (tables/ru-RU/headers-ru-held)", () => {
    expect(local('SUM(Sales[1,5])+Sales[[#This Row],[1,5]]+SUM(Sales[[#Headers],[1,5]])')).toBe('СУММ(Sales[1,5])+Sales[@1,5]+СУММ(Sales[[#Заголовки];[1,5]])')
    expect(local('SUM(Sales[1,5])+Sales[[#This Row],[1,5]]', enUS)).toBe('SUM(Sales[1,5])+Sales[@[1,5]]')
    expect(invariant('=СУММ(Sales[1,5])+Sales[@1,5]+Sales[[#Данные];[a;b]]')).toBe('SUM(Sales[1,5])+Sales[@1,5]+Sales[[#Data],[a;b]]')
  })

  it('a leading = is dropped by toInvariant and kept by toLocal', () => {
    expect(invariant('=1,5')).toBe('1.5')
    expect(local('=SUM(1)')).toBe('=СУММ(1)')
  })
})

describe('errors with positions in the given text, in the language of the locale', () => {
  it('a character that means something else on the other side', () => {
    expect(invariant('СУММ(A1,B1)')).toBe('syntax at 7: Здесь не может стоять «,»')
    expect(invariant('=СУММ(1.5)')).toBe('syntax at 7: Здесь не может стоять «.»')
    expect(local('SUM(1;2)')).toBe('syntax at 5: Здесь не может стоять «;»')
  })

  it('an unterminated string or sheet name, a name inside an array constant', () => {
    expect(invariant('=СУММ("abc;1)')).toBe('syntax at 6: Текст не закрыт кавычкой')
    expect(invariant("'Лист!A1")).toBe('syntax at 0: Имя листа не закрыто апострофом')
    expect(invariant('{1;x}')).toBe('syntax at 3: В массиве констант не может стоять «x»')
  })

  it('text that would read differently once translated', () => {
    // `A1,5` is a reference and a number in ru-RU, but one name `A1.5` in the invariant syntax
    expect(invariant('СУММ(A1,5)')).toBe('syntax at 5: В хранимой формуле «A1» читается иначе')
  })

  it('in English for a locale without messages of its own', () => {
    const own: Locale = { ...ruRU, tag: 'xx-YY' }
    expect(invariant('=СУММ(1.5)', own)).toBe('syntax at 7: unexpected "."')
    expect(invariant('СУММ(A1,5)', own)).toBe('syntax at 5: "A1" cannot be written in the invariant syntax')
    expect(local('SUM(1;2)', own)).toBe('syntax at 5: unexpected ";"')
  })
})

describe('validateLocalFormula', () => {
  it('the result of validateFormula for the invariant text', () => {
    const r = validateLocalFormula('=ОКРУГЛ(СУММ(Данные!A1:A3);2)', ruRU, { sheets: ['Данные'] })
    expect(r).toEqual(validateFormula('ROUND(SUM(Данные!A1:A3),2)', { sheets: ['Данные'] }))
    expect(r.ok && r.canonical).toBe("ROUND(SUM('Данные'!A1:A3),2)")
  })

  it('positions of errors are in the typed text, with or without =', () => {
    const at = (t: string, ctx = {}) => {
      const r = validateLocalFormula(t, ruRU, ctx)
      return r.ok ? 'ok' : `${r.code} at ${r.at}`
    }
    expect(at('=СУММ(A1;НЕТ(1))')).toBe('unknown_function at 9')
    expect(at('СУММ(A1;НЕТ(1))')).toBe('unknown_function at 8')
    expect(at('=ЕСЛИ(1;2;Нет!A1)', { sheets: ['Да'] })).toBe('unknown_sheet at 10')
    expect(at('=СУММ(1,5;')).toBe('syntax at 10')
    expect(at('=ОКРУГЛ(1,5)')).toBe('syntax at 1')
    expect(at('=СУММ(1.5)')).toBe('syntax at 7')
  })
})

describe('locales as data', () => {
  it('en-US is the invariant syntax', () => {
    for (const f of ["SUMIFS('Data'!D:D,'Data'!B:B,\">=1.5\")", 'IF(TRUE,{1,2;3,4},#N/A)', 'ROUND(1.5,0)']) {
      expect(local(f, enUS)).toBe(f)
      expect(invariant(f, enUS)).toBe(f)
    }
  })

  it("a locale of one's own is an object of the same shape", () => {
    const deDE: Locale = {
      tag: 'de-DE',
      decimalSeparator: ',',
      listSeparator: ';',
      arrayColumnSeparator: '.',
      arrayRowSeparator: ';',
      booleans: { TRUE: 'WAHR', FALSE: 'FALSCH' },
      errors: { '#N/A': '#NV' },
      functions: { SUM: 'SUMME', IF: 'WENN' },
    }
    expect(local('IF(TRUE,SUM({1.5,2;3,4}),#N/A)', deDE)).toBe('WENN(WAHR;SUMME({1,5.2;3.4});#NV)')
    expect(invariant('=WENN(WAHR;SUMME({1,5.2;3.4});#NV)', deDE)).toBe('IF(TRUE,SUM({1.5,2;3,4}),#N/A)')
  })

  it('a malformed locale or text is an InputError', () => {
    const bad = (patch: Partial<Record<keyof Locale, unknown>>) => () => toLocal('1', { ...ruRU, ...patch } as Locale)
    expect(bad({ decimalSeparator: ';' })).toThrow('the decimal separator, the list separator and the array separators must differ')
    expect(bad({ listSeparator: 'x' })).toThrow('locale.listSeparator: expected one character that has no other meaning in a formula, got "x"')
    expect(bad({ listSeparator: ':' })).toThrow(InputError)
    expect(bad({ booleans: { TRUE: 'ДА' } })).toThrow('locale.booleans: expected both TRUE and FALSE')
    expect(bad({ errors: { '#OOPS!': '#ОЙ!' } })).toThrow('locale.errors: unknown key "#OOPS!"')
    expect(bad({ functions: { SUM: 'СУММ', SUMSQ: 'сумм' } })).toThrow('locale.functions["SUMSQ"]: "сумм" is also the word for SUM')
    expect(bad({ functions: { SUM: 'С(' } })).toThrow('expected a function name')
    expect(() => toLocal(1 as never, ruRU)).toThrow('formula: expected a string, got number 1')
    expect(() => toInvariant('1', null as never)).toThrow('locale: expected a locale object')
  })
})
