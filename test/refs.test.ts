/** Ссылки: кириллические листы, кавычки в критерии, три способа сослаться на колонку, области имён. */
import { describe, expect, it } from 'vitest'
import { compute, type BookInput } from '../src/index'

const clock = { now: '2026-10-04T12:00:00Z', seed: 1 }
const data = { A1: { value: 'он сказал "да"' }, A2: { value: 'нет' }, A3: { value: 'он сказал "да"' }, B1: { value: 1 }, B2: { value: 2 }, B3: { value: 4 } }

describe('ссылки', () => {
  it('кириллический лист с кавычками и без после канона одно и то же', () => {
    const book: BookInput = { sheets: [{ name: 'Операции', cells: data }, { name: 'S', cells: { A1: { formula: 'SUM(Операции!B:B)' }, A2: { formula: "SUM('Операции'!B:B)" }, A3: { formula: 'SUM(операции!B1:B3)' } } }] }
    const c = compute(book, { clock }).sheets[1].cells
    expect([c.A1.value, c.A2.value, c.A3.value]).toEqual([7, 7, 7])
    expect(c.A1.formula).toBe("SUM('Операции'!B:B)")
    expect(c.A3.formula).toBe("SUM('Операции'!B1:B3)")
  })

  it('критерий с кавычками внутри', () => {
    const book: BookInput = { sheets: [{ name: 'D', cells: data }, { name: 'S', cells: { A1: { formula: 'SUMIFS(D!B:B,D!A:A,"он сказал ""да""")' } } }] }
    expect(compute(book, { clock }).sheets[1].cells.A1.value).toBe(5)
  })

  it('целая колонка, диапазон и имя дают одно', () => {
    const book: BookInput = {
      sheets: [{ name: 'D', cells: data }, { name: 'S', cells: { A1: { formula: 'SUMIFS(D!B:B,D!A:A,"нет")' }, A2: { formula: 'SUMIFS(D!$B$1:$B$3,D!$A$1:$A$3,"нет")' }, A3: { formula: 'SUMIFS(amt,txt,"нет")' } } }],
      names: [{ name: 'amt', ref: "'D'!$B:$B" }, { name: 'txt', ref: "'D'!$A:$A" }],
    }
    const c = compute(book, { clock }).sheets[1].cells
    expect([c.A1.value, c.A2.value, c.A3.value]).toEqual([2, 2, 2])
  })

  it('локальное имя перекрывает глобальное на своём листе', () => {
    const book: BookInput = {
      sheets: [{ name: 'D', cells: data }, { name: 'S', cells: { A1: { formula: 'x*10' } } }, { name: 'T', cells: { A1: { formula: 'x*10' } } }],
      names: [{ name: 'x', ref: "'D'!$B$1" }, { name: 'x', ref: "'D'!$B$3", scope: 'T' }],
    }
    const r = compute(book, { clock })
    expect(r.sheets[1].cells.A1.value).toBe(10)
    expect(r.sheets[2].cells.A1.value).toBe(40)
  })

  it('недопустимые имена отклоняются, формулы с ними дают #NAME?', () => {
    const book: BookInput = {
      sheets: [{ name: 'D', cells: data }, { name: 'S', cells: { A1: { formula: 'SUM(Q1)' }, A2: { formula: 'SUM(Сумма)' } } }],
      names: ['Q1', 'R1C1', 'TRUE', 'Сумма'].map((name) => ({ name, ref: "'D'!$B:$B" })),
    }
    const r = compute(book, { clock })
    expect(r.nameIssues.map((n) => n.name)).toEqual(['Q1', 'R1C1', 'TRUE', 'Сумма'])
    // Q1 это адрес ячейки (пустой), Сумма неизвестное имя
    expect(r.sheets[1].cells.A1.value).toBe(0)
    expect(r.sheets[1].cells.A2.value).toEqual({ error: 'NAME' })
  })
})
