/** Structured references as text: reading every form Excel takes, the stored (canonical) text, the forms Excel refuses. */
import { describe, expect, it } from 'vitest'
import { validateFormula } from '../src/index.js'

const tables = [
  { name: 'Sales', columns: ['Region', 'Qty', 'Price', 'Total'] },
  { name: 'Wierd', columns: ['a b', 'x#y', "p'q", '[br]', 'Total $', '@at', '2024', 'Ünï'] },
]

const canon = (text: string) => {
  const r = validateFormula(text, { tables })
  if (!r.ok) throw new Error(`${text}: ${r.message}`)
  return r.canonical
}

describe('structured references', () => {
  it('store the forms an xlsx file stores (Excel 16.113 for Mac saved these)', () => {
    expect(canon('Sales[@Qty]')).toBe('Sales[[#This Row],[Qty]]')
    expect(canon('Sales[#This Row]')).toBe('Sales[#This Row]')
    expect(canon('Sales[@]')).toBe('Sales[#This Row]')
    expect(canon('Sales[@[Qty]:[Price]]')).toBe('Sales[[#This Row],[Qty]:[Price]]')
    expect(canon('ROWS(Sales)')).toBe('ROWS(Sales[])')
    expect(canon('ROWS(Sales[])')).toBe('ROWS(Sales[])')
    expect(canon('SUM(Sales[[Qty]:[Price]])')).toBe('SUM(Sales[[Qty]:[Price]])')
    expect(canon('SUM(Sales[[Price]:[Qty]])')).toBe('SUM(Sales[[Price]:[Qty]])')
    expect(canon('ROWS(Sales[#All])')).toBe('ROWS(Sales[#All])')
    expect(canon('ROWS(Sales[#Data])')).toBe('ROWS(Sales[#Data])')
    expect(canon('Sales[[#Headers],[Qty]]')).toBe('Sales[[#Headers],[Qty]]')
    expect(canon('SUM(Sales[[#Headers],[#Data],[Qty]])')).toBe('SUM(Sales[[#Headers],[#Data],[Qty]])')
    expect(canon('SUM(Sales[[#Totals],[#Data],[Qty]])')).toBe('SUM(Sales[[#Data],[#Totals],[Qty]])')
  })

  it('take any case, spaces where Excel takes them, and Excel\'s escapes', () => {
    expect(canon('sales[qty]')).toBe('Sales[Qty]')
    expect(canon('SUM(Sales[ [Qty]:[Price] ])')).toBe('SUM(Sales[[Qty]:[Price]])')
    expect(canon('SUM(Sales[[#Data], [Qty]])')).toBe('SUM(Sales[[#Data],[Qty]])')
    expect(canon('SUM(Sales[[#data],[qty]])')).toBe('SUM(Sales[[#Data],[Qty]])')
    expect(canon('SUM(Wierd[a b])')).toBe('SUM(Wierd[a b])')
    expect(canon('SUM(Wierd[[a b]])')).toBe('SUM(Wierd[a b])')
    expect(canon("SUM(Wierd[x'#y])")).toBe("SUM(Wierd[x'#y])")
    expect(canon("SUM(Wierd[p''q])")).toBe("SUM(Wierd[p''q])")
    expect(canon("SUM(Wierd['[br']])")).toBe("SUM(Wierd['[br']])")
    expect(canon('SUM(Wierd[Total $])')).toBe('SUM(Wierd[Total $])')
    expect(canon("SUM(Wierd['@at])")).toBe("SUM(Wierd['@at])")
    expect(canon('SUM(Wierd[2024])')).toBe('SUM(Wierd[2024])')
    expect(canon('SUM(Wierd[Ünï])')).toBe('SUM(Wierd[Ünï])')
    expect(canon("SUM(Wierd[[x'#y]:[p''q]])")).toBe("SUM(Wierd[[x'#y]:[p''q]])")
    expect(canon('SUM(S!Sales[Qty])')).toBe('SUM(Sales[Qty])')
  })

  it('refuse what Excel does not take', () => {
    const bad = (text: string, code: string) => expect(validateFormula(text, { tables })).toMatchObject({ ok: false, code })
    bad('SUM(Sales[[Qty] : [Price]])', 'syntax')
    bad('Sales[[#Headers],[#Totals]]', 'syntax')
    bad('Sales[[#All],[#Data]]', 'syntax')
    bad('Sales[[#This Row],[#Data]]', 'syntax')
    bad('Sales[[Qty],[Price]]', 'syntax')
    bad('Sales[#Nope]', 'syntax')
    bad('SUM(Sales[ Qty ])', 'unknown_name')
    bad('SUM(Sales[Nope])', 'unknown_name')
    bad('SUM(Nope[Qty])', 'unknown_name')
    bad('[@Qty]', 'unknown_name')
    bad('SUM([Qty])', 'unknown_name')
    bad('[1]Sheet1!A1', 'external_ref')
  })

  it('are checked only against tables given; the tables a formula names are listed', () => {
    const r = validateFormula('SUM(Sales[Qty])+[@Price]')
    expect(r).toMatchObject({ ok: true, canonical: 'SUM(Sales[Qty])+[[#This Row],[Price]]', tables: ['Sales'] })
    expect(validateFormula('ROWS(Sales)', { names: [] })).toMatchObject({ ok: false, code: 'unknown_name' })
    expect(validateFormula('ROWS(Sales)', { names: [], tables })).toMatchObject({ ok: true, canonical: 'ROWS(Sales[])', names: [], tables: ['Sales'] })
  })
})
