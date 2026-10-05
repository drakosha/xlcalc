/** Ошибки как значения: код в ячейке, причина в `issues`, неверная формула не роняет расчёт. */
import { describe, expect, it } from 'vitest'
import { compute, type BookInput, type CellInput } from '../src/index.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 1 }
function run(cells: Record<string, CellInput>, extra: BookInput['sheets'] = []) {
  const r = compute({ sheets: [{ name: 'S', cells }, ...extra] }, { clock })
  const v = (a: string) => r.sheets[0].cells[a].value
  const issue = (a: string) => r.issues.find((i) => i.sheet === 'S' && i.cell === a)
  return { r, v, issue }
}

describe('ошибки', () => {
  it('#DIV/0! через SUM и IFERROR, причина с ячейкой-источником', () => {
    const { v, issue } = run({ A1: { value: 1 }, A2: { formula: '1/0' }, A3: { formula: 'SUM(A1:A2)' }, A4: { formula: 'IFERROR(SUM(A1:A2),-1)' }, A5: { formula: 'A3*2' } })
    expect(v('A2')).toEqual({ error: 'DIV0' })
    expect(v('A3')).toEqual({ error: 'DIV0' })
    expect(v('A4')).toBe(-1)
    expect(issue('A2')!.reason).toBe('division by zero')
    expect(issue('A3')!.reason).toBe("division by zero, from 'S'!A2")
    expect(issue('A5')!.reason).toBe("division by zero, from 'S'!A3")
  })

  it('#REF!: неизвестный лист и литерал', () => {
    const { v, issue } = run({ A1: { formula: 'Нет!A1+1' }, A2: { formula: '#REF!' } })
    expect(v('A1')).toEqual({ error: 'REF' })
    expect(issue('A1')).toMatchObject({ check: 'unknown_sheet', reason: 'unknown sheet Нет' })
    expect(v('A2')).toEqual({ error: 'REF' })
  })

  it('#NAME?: неизвестная функция и имя', () => {
    const { v, issue } = run({ A1: { formula: 'VLOOKUP(1,B1:C2,2)' }, A2: { formula: 'nope*2' } })
    expect(v('A1')).toEqual({ error: 'NAME' })
    expect(issue('A1')).toMatchObject({ check: 'unknown_function', at: 0 })
    expect(v('A2')).toEqual({ error: 'NAME' })
    expect(issue('A2')!.check).toBe('unknown_name')
  })

  it('цикл: #CYCLE! у ячеек цикла и зависящих от них, остальное считается', () => {
    const { v, issue } = run({ A1: { formula: 'A2+1' }, A2: { formula: 'A1+1' }, A3: { formula: 'A1*2' }, A4: { formula: '5' }, B1: { formula: 'B1' } })
    expect([v('A1'), v('A2'), v('A3'), v('B1')]).toEqual(Array(4).fill({ error: 'CYCLE' }))
    expect(v('A4')).toBe(5)
    expect(issue('A1')!.reason).toBe('circular reference')
  })

  it('операция над массивом: код проверки unsupported (растекания в этой версии нет)', () => {
    const { v, issue } = run({ A1: { value: 1 }, A2: { value: 2 }, B1: { formula: 'A1:A2*10' } })
    expect(v('B1')).toEqual({ error: 'ERROR' })
    expect(issue('B1')!.check).toBe('unsupported')
  })

  it('неверные формулы книги не роняют расчёт: форма из Python-книги пробы даёт syntax', () => {
    const { v, issue, r } = run({ A1: { formula: 'SUM(B1:B2),SUM(C1:C2)' }, A2: { formula: '' }, A3: { formula: 'A1+1' }, A4: { formula: '2+2' } })
    expect(issue('A1')!.check).toBe('syntax')
    expect(issue('A2')!.check).toBe('empty')
    expect(v('A3')).toEqual({ error: 'ERROR' })
    expect(v('A4')).toBe(4)
    expect(r.issues.map((i) => i.cell)).toEqual(['A1', 'A2', 'A3'])
  })

  it('внешние функции: значение из externals, нет значения это #N/A', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { formula: 'BALANCE("2025-01")*2' }, A2: { formula: 'BALANCE("2025-02")' } } }] }
    const r = compute(book, { clock, externals: { functions: ['BALANCE'], values: { 'BALANCE("2025-01")': 21 } } })
    expect(r.sheets[0].cells.A1.value).toBe(42)
    expect(r.sheets[0].cells.A2.value).toEqual({ error: 'NA' })
  })
})
