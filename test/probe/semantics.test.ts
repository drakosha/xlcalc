/**
 * Семантика списка шага 1 во всех вариантах. Ожидаемое по документации Excel; оракул LibreOffice
 * (если есть `soffice`) подтверждает его, кроме записанных расхождений самого LibreOffice.
 */
import { beforeAll, describe, expect, it } from 'vitest'
import { hasSoffice } from './oracle'
import { runSemantics, sameValue, type SemRow } from './semantics-run'

/** Где вариант расходится с Excel (и почему); всё остальное обязано совпасть. */
const KNOWN: Record<'H' | 'C' | 'Cfj' | 'oracle', Record<string, string>> = {
  H: {
    x1: 'поиск HF снимает диакритику всегда: «Чаи» находит «Чай»',
    r3: 'ROUND(1.005,2) даёт 1: HF округляет двоичное 1.00499…',
    t2: 'TEXT HF не знает формат с разделителем тысяч',
    t3: 'TEXT HF не знает процентный формат',
  },
  C: {
    t4: 'TEXT formulajs не знает форматов даты: серийный номер выходит числом',
    t5: 'TEXT formulajs падает на формате DD.MM.YYYY (адаптер даёт #VALUE!)',
  },
  Cfj: {
    s1: 'formulajs: дата текстом в критерии не сравнивается',
    s3: 'formulajs: нет шаблонов',
    s4: 'formulajs: нет шаблонов',
    s5: 'formulajs: учитывает регистр',
    s7: 'formulajs: учитывает регистр',
    s8: 'formulajs: нет шаблонов',
    c1: 'formulajs: дата текстом в критерии не сравнивается',
    c2: 'formulajs: нет шаблонов',
    c3: 'formulajs: учитывает регистр',
    t4: 'как у C',
    t5: 'как у C',
  },
  oracle: {
    d5: 'LibreOffice не повторяет ложное 29.02.1900: до 01.03.1900 серийный номер на 1 больше',
    u1: 'в LibreOffice логическое это число: TRUE в диапазоне считается',
    u2: 'LibreOffice по умолчанию не приводит текст аргумента к числу (#VALUE!)',
  },
}

/** Функции списка семантики, где голая formulajs неверна, и лечит ли это наш адаптер аргументов. */
const FJ_WRONG = {
  notFixable: ['COUNTIFS', 'SUMIFS', 'TEXT'],
  fixedByAdapter: ['DATE', 'EOMONTH', 'IF', 'IFERROR', 'SUM'],
  missing: ['XLOOKUP'],
}

let rows: SemRow[]
beforeAll(async () => {
  rows = await runSemantics(hasSoffice())
}, 120_000)

describe('семантика списка шага 1', () => {
  for (const variant of ['H', 'C', 'Cfj'] as const) {
    it(`${variant}: совпадает с Excel, кроме записанных расхождений`, () => {
      const off = rows.filter((r) => !sameValue(r[variant], r.c.excel)).map((r) => r.c.id)
      expect(off.sort()).toEqual(Object.keys(KNOWN[variant]).sort())
    })
  }

  it.skipIf(!hasSoffice())('оракул LibreOffice подтверждает ожидаемое, кроме своих расхождений', () => {
    const off = rows.filter((r) => !sameValue(r.oracle, r.c.excel)).map((r) => r.c.id)
    expect(off.sort()).toEqual(Object.keys(KNOWN.oracle).sort())
  })

  it('голая formulajs: неверные функции и что лечит адаптер', () => {
    const wrong = new Set(rows.filter((r) => !sameValue(r.fjRaw, r.c.excel)).map((r) => r.c.fn))
    const missing = [...wrong].filter((fn) => rows.some((r) => r.c.fn === fn && r.fjRaw === 'нет функции'))
    // не лечится: и с адаптером (вариант C со SUMIFS formulajs) ответ неверный
    const notFixable = [...wrong].filter((fn) => !missing.includes(fn) && rows.some((r) => r.c.fn === fn && !sameValue(r.Cfj, r.c.excel)))
    const fixed = [...wrong].filter((fn) => !missing.includes(fn) && !notFixable.includes(fn))
    expect(missing.sort()).toEqual(FJ_WRONG.missing)
    expect(notFixable.sort()).toEqual(FJ_WRONG.notFixable)
    expect(fixed.sort()).toEqual(FJ_WRONG.fixedByAdapter)
    // стоп-правило хендоффа: больше трёх неверных по оракулу и не лечимых адаптером
    expect(notFixable.length).toBeLessThanOrEqual(3)
  })
})
