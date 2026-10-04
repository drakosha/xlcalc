/**
 * Особенности HyperFormula 3.4.0, найденные в исходниках (`docs/design-workbook.md`, восемь пунктов):
 * каждая проверяется на голом HF и на нашем адаптере (канонизатор и замороженная конфигурация).
 * Это тест конкретного движка, а не контракта: при смене движка он уходит.
 */
import { HyperFormula, type RawCellContent } from 'hyperformula'
import { describe, expect, it } from 'vitest'
import { hfEngine, HF_CONFIG } from '../../src/probe/hf-engine'
import { parse, toHfText } from '../../src/probe/parser'
import type { BookInput, CellValue } from '../../src/probe/model'

type Raw = Record<string, RawCellContent[][]>

/** Значения первого листа голого HF; ошибка как `#ТИП`. */
function raw(sheets: Raw, config: Record<string, unknown> = {}, cells: [number, number][] = [[0, 0]]): unknown[] {
  const hf = HyperFormula.buildFromSheets(sheets, { licenseKey: 'gpl-v3', ...config })
  const out = cells.map(([row, col]) => {
    const v = hf.getCellValue({ sheet: 0, row, col }) as unknown
    return v !== null && typeof v === 'object' ? `#${(v as { type: string }).type}` : v
  })
  hf.destroy()
  return out
}

const clock = { now: '2026-10-04T12:00:00Z', seed: 1 }
function adapter(book: BookInput, sheet: string, addr: string): CellValue | undefined {
  return hfEngine.compute(book, { clock }).values.get(sheet)?.get(addr)
}

describe('HF 3.4.0: особенности', () => {
  it('1. лексер не знает кириллицу: лист без кавычек не разбирается, имя диапазона кириллицей недопустимо', () => {
    const hf = HyperFormula.buildFromSheets({ Операции: [[1], [2]], S: [['=SUM(Операции!A:A)'], ["=SUM('Операции'!A:A)"]] }, { licenseKey: 'gpl-v3' })
    const s = hf.getSheetId('S')!
    expect((hf.getCellValue({ sheet: s, row: 0, col: 0 }) as { type: string }).type).toBe('ERROR')
    expect(hf.getCellValue({ sheet: s, row: 1, col: 0 })).toBe(3)
    hf.destroy()
    expect(() => HyperFormula.buildFromSheets({ S: [[1]] }, { licenseKey: 'gpl-v3' }, [{ name: 'Сумма', expression: '=S!$A$1' }])).toThrow(/invalid/)
    // адаптер: канон всегда ставит лист в кавычки
    expect(toHfText(parse('SUM(Операции!A:A)'))).toBe("=SUM('Операции'!A:A)")
    const book: BookInput = { sheets: [{ name: 'Операции', cells: { A1: { value: 1 }, A2: { value: 2 } } }, { name: 'S', cells: { A1: { formula: 'SUM(Операции!A:A)' } } }] }
    expect(adapter(book, 'S', 'A1')).toBe(3)
  })

  it('2. строковый литерал без удвоения "": кавычка и обратная косая через CHAR(34)/CHAR(92)', () => {
    expect(raw({ S: [['="a""b"'], ['="a\\"'], ['="a\\b"']] }, {}, [[0, 0], [1, 0], [2, 0]])).toEqual(['#ERROR', '#ERROR', 'a\\b'])
    expect(toHfText(parse('"он сказал ""да"""'))).toBe('=("он сказал "&CHAR(34)&"да"&CHAR(34))')
    expect(toHfText(parse('"путь\\"'))).toBe('=("путь"&CHAR(92))')
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { formula: '"он сказал ""да"""' }, A2: { formula: '"путь\\"' } } }] }
    expect(adapter(book, 'S', 'A1')).toBe('он сказал "да"')
    expect(adapter(book, 'S', 'A2')).toBe('путь\\')
  })

  it('3. выдача округляется до precisionRounding (по умолчанию 10): у сумм от 1 млрд теряются копейки; -0 HF не отдаёт', () => {
    const sheets: Raw = { S: [[1234567890.12, '=A1*1', '=0.1+0.2', 0, '=-D1', '=ROUND(-0.001,2)', -0]] }
    const cells: [number, number][] = [[0, 1], [0, 2], [0, 4], [0, 5], [0, 6]]
    expect(raw(sheets, {}, cells).slice(0, 2)).toEqual([1234567890.1, 0.3])
    expect(raw(sheets, { precisionRounding: 14, smartRounding: true }, cells).slice(0, 2)).toEqual([1234567890.12, 0.3])
    // в 3.4.0 -0 не проходит ни из значения, ни из формулы (пункт хендоффа не подтвердился); канон всё равно пишет 0
    for (const v of raw(sheets, { precisionRounding: 14 }, cells).slice(2)) expect(Object.is(v, -0)).toBe(false)
  })

  it('4. accentSensitive:false склеивает «й» и «и»; поиск (XLOOKUP, MATCH) склеивает всегда', () => {
    const sheets: Raw = { S: [['и'], ['й'], ['=COUNTIF(A1:A2,"й")'], ['=XLOOKUP("й",A1:A2,A1:A2)'], ['=MATCH("й",A1:A2,0)']] }
    const cells: [number, number][] = [[2, 0], [3, 0], [4, 0]]
    expect(raw(sheets, { accentSensitive: false }, cells)).toEqual([2, 'и', 1])
    // с нашей конфигурацией критерии различают, а поиск нет: forceNormalizeString снимает диакритику без учёта настройки
    expect(raw(sheets, { accentSensitive: true }, cells)).toEqual([1, 'и', 1])
  })

  it('5. TRUE и FALSE не ключевые слова: без скобок #NAME?, канон пишет вызов', () => {
    expect(raw({ S: [['=TRUE'], ['=TRUE()']] }, {}, [[0, 0], [1, 0]])).toEqual(['#NAME', true])
    expect(toHfText(parse('IF(TRUE,1,2)'))).toBe('=IF(TRUE(),1,2)')
    expect(adapter({ sheets: [{ name: 'S', cells: { A1: { formula: 'IF(FALSE,1,2)' } } }] }, 'S', 'A1')).toBe(2)
  })

  it('6. HF угадывает тип строки: строки идут с апострофом, даты серийным номером', () => {
    const sheets: Raw = { S: [['TRUE', '25%', '=1+1', '2025-01-01', '01/02/2025', "'TRUE"]] }
    const cells: [number, number][] = [0, 1, 2, 3, 4, 5].map((c) => [0, c])
    expect(raw(sheets, {}, cells)).toEqual([true, 0.25, 2, '2025-01-01', 45689, 'TRUE'])
    expect(raw(sheets, { dateFormats: ['YYYY-MM-DD'] }, cells)).toEqual([true, 0.25, 2, 45658, '01/02/2025', 'TRUE'])
    const cellsIn = { A1: { value: 'TRUE' }, A2: { value: '25%' }, A3: { value: '=1+1' }, A4: { value: '2025-01-01' }, A5: { value: "'x" }, A6: { value: { date: '2025-01-01' } } }
    const v = hfEngine.compute({ sheets: [{ name: 'S', cells: cellsIn }] }, { clock }).values.get('S')!
    expect([...v.values()]).toEqual(['TRUE', '25%', '=1+1', '2025-01-01', "'x", 45658])
  })

  it('7. useArrayArithmetic:false: диапазон в скалярном месте даёт неявное пересечение по строке, вне строк диапазона #VALUE! (пункт хендоффа о сворачивании к левому верхнему не подтвердился)', () => {
    const off: Raw = { S: [[1, '=ABS(A1:A2)', '="a"&A1:A2'], [2, '=ABS(A1:A2)', '="a"&A1:A2'], [null], [null, '=ABS(A1:A2)', '="a"&A1:A2']] }
    expect(raw(off, { useArrayArithmetic: false }, [[0, 1], [1, 1], [3, 1], [0, 2], [1, 2], [3, 2]])).toEqual([1, 2, '#VALUE', 'a1', 'a2', '#VALUE'])
    // у нас true: результат растекается вниз, как в Excel 365
    const on: Raw = { S: [[1], [2], [null], [null, '=ABS(A1:A2)', '=A1:A2*10']] }
    expect(raw(on, { useArrayArithmetic: true }, [[3, 1], [4, 1], [3, 2], [4, 2]])).toEqual([1, 2, 10, 20])
    expect(HF_CONFIG.useArrayArithmetic).toBe(true)
  })

  it('8. встроенные TODAY, NOW, RAND, RANDBETWEEN заменяются своими (registerFunction, момент и зерно из config.context)', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { formula: 'TODAY()' }, A2: { formula: 'NOW()' }, A3: { formula: 'RAND()' }, A4: { formula: 'RANDBETWEEN(1,1000000)' } } }] }
    const a = hfEngine.compute(book, { clock: { now: '2025-01-01T06:00:00Z', seed: 42 } })
    const b = hfEngine.compute(book, { clock: { now: '2025-01-01T06:00:00Z', seed: 42 } })
    const c = hfEngine.compute(book, { clock: { now: '2025-01-01T06:00:00Z', seed: 43 } })
    const s = a.values.get('S')!
    expect(s.get('A1')).toBe(45658)
    expect(s.get('A2')).toBe(45658.25)
    expect(a.hash).toBe(b.hash)
    expect(c.values.get('S')!.get('A3')).not.toBe(s.get('A3'))
    expect(a.volatileCells).toEqual(['S!A1', 'S!A2', 'S!A3', 'S!A4'])
  })
})
