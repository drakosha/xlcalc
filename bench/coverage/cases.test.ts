/**
 * Генератор случаев замера: страж размеров массивов по всем случаям и полнота набора.
 */
import { describe, expect, it } from 'vitest'
import { arity, buildCases, parseKinds } from './cases.js'
import { ARRAY_SIZE, SIGNATURES, SIZE_LIMIT } from './signatures.js'

const set = buildCases()

describe('случаи замера покрытия', () => {
  it('в позиции размера массива только маленькие числовые литералы', () => {
    const bad: string[] = []
    for (const c of set.cases) {
      for (const i of ARRAY_SIZE[c.fn] ?? []) {
        if (i >= c.args.length) continue
        const a = c.args[i]
        if (typeof a !== 'number' || !Number.isFinite(a) || Math.abs(a) > SIZE_LIMIT) bad.push(c.formula)
      }
    }
    expect(bad).toEqual([])
    // в книге аргумент-размер так и записан литералом
    for (const c of set.cases.filter((x) => x.fn in ARRAY_SIZE)) expect(c.formula).not.toMatch(/MUNIT\((?!-?[\d.]+\))/)
  })

  it('функции-массивы с целыми аргументами разобраны: размер или индекс', () => {
    // CHOOSECOLS, CHOOSEROWS и SORT берут номера колонок и строк: массив не больше входного
    const index = ['CHOOSECOLS', 'CHOOSEROWS', 'SORT']
    const withInt = SIGNATURES.filter((s) => s.picks && parseKinds(s).some((k) => k.k === 'i')).map((s) => s.name)
    expect(withInt.sort()).toEqual([...Object.keys(ARRAY_SIZE), ...index].sort())
  })

  it('число аргументов в пределах Excel, на функцию от 1 до 25 случаев', () => {
    const per = new Map<string, number>()
    for (const c of set.cases) per.set(c.fn, (per.get(c.fn) ?? 0) + 1)
    expect(per.size).toBe(402)
    for (const n of per.values()) expect(n).toBeGreaterThan(0)
    for (const n of per.values()) expect(n).toBeLessThanOrEqual(25)
    for (const s of SIGNATURES) {
      const { min, max } = arity(s)
      for (const c of set.cases.filter((x) => x.fn === s.name)) expect(c.args.length >= min && c.args.length <= max).toBe(true)
    }
  })

  it('детерминизм: тот же сид даёт те же случаи', () => {
    expect(buildCases().cases).toEqual(set.cases)
  })
})
