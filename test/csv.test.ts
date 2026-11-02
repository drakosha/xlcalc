/**
 * The `./csv` subpath without Excel: `readCsv` takes a file's bytes or text, finds its encoding and delimiter, reads its
 * fields as RFC 4180 and each as pasted in the locale, keeps formulas and codes as text when asked, refuses past its
 * limit and what it does not take; `writeCsv` writes a sheet as Excel saves it, in UTF-8, UTF-16LE or the code page.
 * Excel's own answers are in `csv-excel.test.ts`.
 */
import { describe, expect, it } from 'vitest'
import { readCsv, writeCsv } from '../src/csv.js'
import { compute, createSession, InputError, registerLocale } from '../src/index.js'
import { ruRU } from '../src/locales/ru-RU.js'

registerLocale(ruRU)
const now = '2026-10-10T12:00:00Z'
const bytes = (...b: number[]) => new Uint8Array(b)
const utf8 = (s: string) => new TextEncoder().encode(s)
const cells = (data: Uint8Array | string, opts = {}) => readCsv(data, { now, ...opts }).book.sheets[0].cells

describe('readCsv: encodings', () => {
  it('a byte order mark, UTF-16 by its zeros, UTF-8 when valid, else the locale\'s code page', () => {
    expect(readCsv(bytes(0xef, 0xbb, 0xbf, 0x61)).encoding).toBe('utf-8')
    expect(readCsv(bytes(0xff, 0xfe, 0x61, 0, 0x2c, 0, 0x62, 0)).encoding).toBe('utf-16le')
    expect(readCsv(bytes(0xfe, 0xff, 0, 0x61)).encoding).toBe('utf-16be')
    expect(readCsv(bytes(0x61, 0, 0x2c, 0, 0x62, 0)).encoding).toBe('utf-16le')
    expect(readCsv(bytes(0, 0x61, 0, 0x2c, 0, 0x62)).encoding).toBe('utf-16be')
    expect(readCsv(utf8('Café,1')).encoding).toBe('utf-8')
    expect(cells(bytes(0x43, 0x61, 0x66, 0xe9))).toEqual({ A1: { value: 'Café' } })
    // € and ™ by the code page's table: Node 20's TextDecoder reads windows-1252 as Latin-1
    expect([cells(bytes(0x80, 0x2c, 0x99)), cells(bytes(0x80), { encoding: 'windows-1252' })]).toEqual([{ A1: { value: '€' }, B1: { value: '™' } }, { A1: { value: '€' } }])
    const ru = readCsv(bytes(0xcf, 0xf0, 0xe8, 0xe2, 0xe5, 0xf2), { locale: ruRU })
    expect([ru.encoding, ru.book.sheets[0].cells.A1, ru.book.locale]).toEqual(['windows-1251', { value: 'Привет' }, 'ru-RU'])
    expect(readCsv(bytes(0xe9), { encoding: 'latin1' }).encoding).toBe('windows-1252')
    expect(readCsv('﻿a,b').encoding).toBeNull()
    expect(cells('﻿a,b')).toEqual({ A1: { value: 'a' }, B1: { value: 'b' } })
    expect(() => readCsv('a', { encoding: 'utf-9' })).toThrow(new InputError('opts.encoding', 'expected an encoding\'s label such as "utf-8", "utf-16le" or "windows-1251", got "utf-9"'))
  })

  it('a file marked as UTF-16 is read at tabs as Excel reads it; the guess takes the delimiter that splits it evenly', () => {
    const file = bytes(0xff, 0xfe, ...[...'a,b\tc'].flatMap((ch) => [ch.charCodeAt(0), 0]))
    expect(readCsv(file, { delimiter: 'locale' }).delimiter).toBe('\t')
    expect(readCsv(file).delimiter).toBe(',')
  })
})

describe('readCsv: delimiters and fields', () => {
  it('sep= names the delimiter and is no row, a host\'s wins; .tsv is tabs; else the guess, ; before , on a tie', () => {
    expect(readCsv('sep=;\r\na;b,c').delimiter).toBe(';')
    expect(cells('SEP=;\na;b')).toEqual({ A1: { value: 'a' }, B1: { value: 'b' } })
    expect(cells('sep=;\na;b,c', { delimiter: ',' })).toEqual({ A1: { value: 'a;b' }, B1: { value: 'c' } })
    expect(readCsv('a\tb,c', { name: 'data.TSV' })).toMatchObject({ delimiter: '\t', book: { sheets: [{ name: 'data' }] } })
    expect(readCsv('1,5;2,5\n3,5;4').delimiter).toBe(';')
    expect(readCsv('a,b\nc,d\ne').delimiter).toBe(',')
    expect(readCsv('a\tb\nc\td').delimiter).toBe('\t')
    expect(readCsv('one\ntwo').delimiter).toBe(',')
    expect(readCsv('one\ntwo', { locale: ruRU }).delimiter).toBe(';')
    expect(readCsv('a;b', { delimiter: 'locale' }).delimiter).toBe(',')
    expect(readCsv('a,b', { delimiter: 'locale', locale: ruRU }).delimiter).toBe(';')
    for (const d of ['"', '\n', ';;', 1]) expect(() => readCsv('a', { delimiter: d as string }), String(d)).toThrow(InputError)
  })

  it('quotes, doubled quotes, line breaks in quotes, CRLF, LF and CR, no row after the last line break', () => {
    expect(cells('"a,b","say ""hi""","x\ny"\r\nc\rd\n')).toEqual({ A1: { value: 'a,b' }, B1: { value: 'say "hi"' }, C1: { value: 'x\ny', style: 0 }, A2: { value: 'c' }, A3: { value: 'd' } })
    expect(readCsv('"x\ny",z').book.styles).toEqual([{ alignment: { wrap: true } }])
    expect(readCsv('"x\ry"').book.styles).toBeUndefined()
    expect(cells('a,,c\n\n,,\nd')).toEqual({ A1: { value: 'a' }, C1: { value: 'c' }, A4: { value: 'd' } })
    expect(readCsv('').book.sheets[0].cells).toEqual({})
  })

  it('the sheet is named after the file, repaired as Excel takes a name', () => {
    expect(readCsv('a', { name: '/home/me/Sales 2026.csv' }).book.sheets[0].name).toBe('Sales 2026')
    expect(readCsv('a', { name: 'a[1]:b.csv' }).book.sheets[0].name).toBe('a_1__b')
    expect(readCsv('a', { name: 'x'.repeat(40) + '.csv' }).book.sheets[0].name).toBe('x'.repeat(31))
    expect(readCsv('a').book.sheets[0].name).toBe('Sheet1')
  })
})

describe('readCsv: values', () => {
  it('each field as pasted in the locale: numbers, dates, percentages, money, booleans, errors, an apostrophe kept', () => {
    expect(cells('1234.5,"1,234",5%,$1.50,3/15/2026,TRUE,#N/A,\'007')).toEqual({
      A1: { value: 1234.5 },
      B1: { value: 1234, numberFormat: 3 },
      C1: { value: 0.05, numberFormat: 9 },
      D1: { value: 1.5, numberFormat: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)' },
      E1: { value: 46096, numberFormat: 14 },
      F1: { value: true },
      G1: { formula: '#N/A' },
      H1: { value: "'007" },
    })
    expect(cells('1 234,5;15.03.2026;ИСТИНА;15.3', { locale: ruRU })).toEqual({ A1: { value: 1234.5, numberFormat: 4 }, B1: { value: 46096, numberFormat: 14 }, C1: { value: true }, D1: { value: 46096, numberFormat: 16 } })
  })

  it('formulas are text and an issue unless formulas: true, as Excel takes them then, of the old kind', () => {
    const text = readCsv('=1+2,+A1,-A1,@SUM(1),-5,=,@A1', { now })
    expect(text.book.sheets[0].cells).toEqual({ A1: { value: '=1+2' }, B1: { value: '+A1' }, C1: { value: '-A1' }, D1: { value: '@SUM(1)' }, E1: { value: -5 }, F1: { value: '=' }, G1: { value: '@A1' } })
    expect(text.issues).toEqual([{ code: 'formula', message: 'fields Excel reads as formulas are kept as text (readCsv takes them as formulas with formulas: true)', sheet: 'Sheet1', cells: ['A1', 'B1', 'C1', 'D1'] }])
    const taken = readCsv('=1+2,+A1,-A1,@SUM(1),=A1:A2,=1+,"=СУММ(1;2)"', { now, formulas: true, delimiter: ',' })
    expect(taken.book.sheets[0].cells).toEqual({ A1: { formula: '1+2' }, B1: { formula: '+A1' }, C1: { formula: '-A1' }, D1: { formula: 'SUM(1)' }, E1: { formula: '@A1:A2' }, F1: { value: '=1+' }, G1: { value: '=СУММ(1;2)' } })
    expect(taken.issues.map((i) => [i.code, i.cells])).toEqual([['formula', ['F1', 'G1']]])
    expect(cells('"=СУММ(1;2)";=1,5*2', { locale: ruRU, formulas: true })).toEqual({ A1: { formula: 'SUM(1,2)' }, B1: { formula: '1.5*2' } })
  })

  it('codes stay text with codesAsText: a leading zero, more than 15 digits', () => {
    const row = '00123,007,0,0.5,1234567890123456789,123456789012345,-007'
    expect(cells(row)).toMatchObject({ A1: { value: 123 }, B1: { value: 7 }, E1: { value: 1234567890123450000 } })
    expect(cells(row, { codesAsText: true })).toEqual({ A1: { value: '00123' }, B1: { value: '007' }, C1: { value: 0 }, D1: { value: 0.5 }, E1: { value: '1234567890123456789' }, F1: { value: 123456789012345 }, G1: { value: -7 } })
  })

  it('a field past 32,767 characters is cut, rows past the sheet left out, each with an issue; past limits.cells refused', () => {
    const long = readCsv(`${'x'.repeat(32_800)},b`)
    expect([(long.book.sheets[0].cells.A1.value as string).length, long.issues.map((i) => [i.code, i.cells])]).toEqual([32_767, [['textLength', ['A1']]]])
    const wide = readCsv(','.repeat(16_384) + 'z')
    expect([wide.book.sheets[0].cells, wide.issues.map((i) => i.code)]).toEqual([{}, ['size']])
    expect(() => readCsv('a,b,c\nd', { limits: { cells: 3 } })).toThrow('csv: the file has more than 3 cells that are not empty (limits.cells)')
    expect(Object.keys(cells('a,,c\n', { limits: { cells: 2 } }))).toEqual(['A1', 'C1'])
    expect(() => readCsv('a', { limits: { cells: 0 } })).toThrow(new InputError('opts.limits.cells', 'expected an integer from 1 to 9007199254740991, got number 0'))
    expect(() => readCsv('a', { limits: { bytes: 1 } as never })).toThrow('opts.limits: unknown field "bytes" (expected cells)')
  })

  it('options it does not take are refused; issues in the locale; the book computes', () => {
    expect(() => readCsv('a', { values: 'omit' } as never)).toThrow('opts: unknown field "values" (expected locale, name, delimiter, encoding, formulas, codesAsText, now, limits) (values is an option of writeXlsx)')
    expect(() => readCsv('a', { bom: false } as never)).toThrow('(bom is an option of writeCsv)')
    expect(() => readCsv('a', { formulas: 1 } as never)).toThrow(new InputError('opts.formulas', 'expected a boolean, got number 1'))
    expect(() => readCsv(5 as never)).toThrow(InputError)
    expect(() => readCsv('a', { now: 'soon' })).toThrow(InputError)
    expect(readCsv('=1', { locale: 'ru-RU' }).issues[0].message).toBe('Поля, которые Excel читает как формулы, оставлены текстом (readCsv берёт их формулами с formulas: true)')
    const { book } = readCsv('2;=A1*1,5', { locale: 'ru-RU', formulas: true })
    expect(compute(book).sheets[0].cells.B1.value).toBe(3)
    expect(readCsv(utf8('a,b').buffer as ArrayBuffer).book.sheets[0].cells).toEqual({ A1: { value: 'a' }, B1: { value: 'b' } })
  })
})

describe('writeCsv', () => {
  const book = { sheets: [{ name: 'S', cells: { B2: { value: 1234.5, numberFormat: '#,##0.00' }, C2: { value: 46096, numberFormat: 'dddd, mmmm d, yyyy' }, B3: { value: 'a,"b"' }, C3: { formula: 'B2*2' }, E4: { numberFormat: '0.00' } } }] }

  it('the used range as shown in no column\'s room, quoted where needed, rows joined by CRLF, UTF-8 with its mark', () => {
    const w = writeCsv({ book, result: compute(book) })
    expect(w.text).toBe('"1,234.50","Sunday, March 15, 2026",,\r\n"a,""b""",2469,,\r\n,,,')
    expect([...w.data.subarray(0, 4)]).toEqual([0xef, 0xbb, 0xbf, 0x22])
    expect(w.issues).toEqual([])
    expect(writeCsv({ book: { sheets: [{ name: 'S', cells: {} }] } }).text).toBe('')
    expect(writeCsv({ book: { sheets: [{ name: 'S', cells: { A1: { value: 'a\tb' }, B1: { value: { date: '2026-03-15' }, numberFormat: 14 } } }] } }).text).toBe('"a\tb",3/15/2026')
  })

  it('the delimiter and the shown values of the locale, the code page, UTF-16LE, no mark', () => {
    const ru = { locale: 'ru-RU', sheets: [{ name: 'Лист', cells: { A1: { value: 1234.5, numberFormat: 4 }, B1: { value: true }, C1: { value: 'Привет €' }, D1: { value: 'a;b' } } }] }
    expect(writeCsv({ book: ru }).text).toBe('1 234,50;ИСТИНА;Привет €;"a;b"')
    expect(writeCsv({ book: ru }, { delimiter: ',', locale: 'en-US' }).text).toBe('"1,234.50",TRUE,Привет €,a;b')
    const ansi = writeCsv({ book: ru }, { encoding: 'ansi' })
    expect([ansi.data.length, new TextDecoder('windows-1251').decode(ansi.data), ansi.issues]).toEqual([ansi.text.length, ansi.text, []])
    const lost = writeCsv({ book: ru }, { encoding: 'ansi', locale: 'en-US' })
    expect(lost.issues).toEqual([{ code: 'encoding', message: 'characters the code page windows-1252 lacks are written as ?', sheet: 'Лист' }])
    expect(cells(lost.data, { encoding: 'windows-1252', delimiter: ',' }).C1).toEqual({ value: '?????? €' })
    expect([...writeCsv({ book: ru }, { encoding: 'utf-16le' }).data.subarray(0, 4)]).toEqual([0xff, 0xfe, 0x31, 0])
    expect([...writeCsv({ book: ru }, { bom: false }).data.subarray(0, 1)]).toEqual([0x31])
  })

  it('the active sheet or the one named, with an issue for the others; a session\'s snapshot; formulas need their result', () => {
    const two = { activeSheet: 'B', sheets: [{ name: 'A', cells: { A1: { value: 1 } } }, { name: 'B', cells: { A1: { formula: 'A!A1+1' } } }] }
    const session = createSession(two)
    expect(writeCsv(session.getSnapshot())).toMatchObject({ text: '2', issues: [{ code: 'sheets', message: 'the file holds the sheet "B" alone: the other sheets are left out', sheet: 'B' }] })
    expect(writeCsv(session.getSnapshot(), { sheet: 'a' }).text).toBe('1')
    session.close()
    expect(() => writeCsv({ book: two })).toThrow(new InputError('snapshot.result', "the formula of 'B'!A1 has no value: pass { book, result } with the result of this workbook"))
    expect(writeCsv({ book: two }, { sheet: 'A' }).text).toBe('1')
    expect(() => writeCsv({ book: two }, { sheet: 'C' })).toThrow(new InputError('opts.sheet', 'the workbook has no sheet "C"'))
    expect(() => writeCsv({ book: two }, { encoding: 'utf-32' as never })).toThrow(InputError)
    expect(() => writeCsv({ book: two }, { mode: 'values' } as never)).toThrow('(mode is an option of writeXlsx)')
    const spill = { sheets: [{ name: 'S', cells: { A1: { formula: 'SEQUENCE(2)' } } }] }
    expect(writeCsv({ book: spill, result: compute(spill) }).text).toBe('1\r\n2')
  })

  it('what readCsv reads, writeCsv writes back', () => {
    const text = 'Item,Qty,Date\r\n"a,b",3,3/15/2026\r\n"say ""hi""",1.5,'
    const { book } = readCsv(text, { now })
    expect(writeCsv({ book }, { bom: false }).text).toBe(text)
  })
})
