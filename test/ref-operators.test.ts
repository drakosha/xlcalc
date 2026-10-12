/**
 * The reference operators: the intersection (whitespace between references) and the range operator between
 * references, how formula text reads them (spaces that are no operator, sheets named like cells), their canonical
 * text, and what they answer in both execution modes, in a session and after a structure operation.
 */
import { describe, expect, it } from 'vitest'
import { compute, createSession, editStructure, toLocal, validateFormula, type BookInput } from '../src/index.js'
import { ruRU } from './locales.js'

const canonical = (f: string) => {
  const v = validateFormula(f)
  return v.ok ? v.canonical : `${v.code}: ${v.message}`
}

const cells = { A1: { value: 1 }, B2: { value: 5 }, C3: { value: 7 } }
const values = (formulas: Record<string, string>, more: BookInput['sheets'] = []) => {
  const book: BookInput = { sheets: [{ name: 'S', cells: { ...cells, ...Object.fromEntries(Object.entries(formulas).map(([a, f]) => [a, { formula: f }])) } }, ...more] }
  const out = [compute(book), compute(book, { execution: 'eval' })]
  expect(out[0].hash).toBe(out[1].hash)
  return Object.fromEntries(Object.keys(formulas).map((a) => [a, out[0].sheets[0].cells[a].value]))
}

describe('the reference operators in formula text', () => {
  it('whitespace between references is the intersection, a space in the canonical text; elsewhere it goes', () => {
    expect(canonical('A1:C3  B2:D4')).toBe('A1:C3 B2:D4')
    expect(canonical('A1:C3\nB2:D4')).toBe('A1:C3 B2:D4')
    expect(canonical('(A1,B1) A1:B1')).toBe('(A1,B1) A1:B1')
    expect(canonical('A1:C3 INDEX(B1:D4,0,2)')).toBe('A1:C3 INDEX(B1:D4,0,2)')
    expect(canonical('SUM( A1:A3 ,B1)')).toBe('SUM(A1:A3,B1)')
    expect(canonical('A1 + B1')).toBe('A1+B1')
    expect(canonical('SUM (A1:A3)')).toBe('SUM(A1:A3)')
    expect(canonical('{TRUE, FALSE}')).toBe('{TRUE,FALSE}')
  })

  it('the range operator joins references that are not one range: calls, sheets, a third corner', () => {
    expect(canonical('A1:INDEX(B1:B3,2)')).toBe('A1:INDEX(B1:B3,2)')
    expect(canonical('INDEX(B1:B3,2):A1')).toBe('INDEX(B1:B3,2):A1')
    expect(canonical('S!A1:S!B2')).toBe("'S'!A1:'S'!B2")
    expect(canonical('A1:B2:C3')).toBe('A1:B2:C3')
  })

  it('two sheets named like cells are no reference across sheets: S1:S3!A1 is S1 to A1 on S3', () => {
    expect(canonical('S1:S3!A1')).toBe("S1:'S3'!A1")
    expect(canonical("'S1:S3'!A1")).toBe("'S1:S3'!A1")
    expect(canonical('S1!A1')).toBe("'S1'!A1")
    expect(canonical('Jan:Mar!A1')).toBe("'Jan:Mar'!A1")
    expect(canonical('XFE1:XFE2!A1')).toBe("'XFE1:XFE2'!A1")
  })

  it('locales keep the operators as they are', () => {
    expect(toLocal('SUM(A1:C3 B2:D4,A1:INDEX(B1:B3,2))', ruRU)).toEqual({ ok: true, formula: 'СУММ(A1:C3 B2:D4;A1:ИНДЕКС(B1:B3;2))' })
  })
})

describe('the reference operators computed', () => {
  it('the intersection: a reference to the cells both meet, #NULL! where none', () => {
    const v = values({ E1: 'SUM(A1:C3 B2:D4)', E2: 'A1:C3 B2:B2', E3: 'A1:A3 C1:C3', E4: 'ERROR.TYPE(A1:A3 C1:C3)', E5: 'ISREF(A1:C3 B2:D4)', E6: '-A1:C3 C3:C3', E7: 'SUM((A1:A3,C1:C3) A3:C3)' })
    expect(v).toEqual({ E1: 12, E2: 5, E3: { error: 'NULL' }, E4: 1, E5: true, E6: -7, E7: 7 })
  })

  it('the range operator: the rectangle around both; references on two sheets, values and references across sheets are #VALUE!', () => {
    const v = values({ E1: 'SUM(A1:INDEX(A1:C3,3,3))', E2: 'ROWS(INDEX(A1:A9,2):A5)', E3: 'SUM(A1:T!A1)', E4: 'SUM(A1:B2:C3)', E5: 'SUM(T:T!A1 A1)' }, [{ name: 'T', cells: { A1: { value: 100 } } }])
    expect(v).toEqual({ E1: 13, E2: 4, E3: { error: 'VALUE' }, E4: 13, E5: { error: 'VALUE' } })
  })

  it('S1:S3!A1 is #VALUE! on another sheet and the range A1:S1 on S3', () => {
    const sheet = (name: string, a1: number) => ({ name, cells: { A1: { value: a1 }, B1: { value: 10 * a1 }, C5: { formula: 'SUM(S1:S3!A1)' } } })
    const r = compute({ sheets: [sheet('S1', 1), sheet('S2', 4), sheet('S3', 7), { name: 'C', cells: { A1: { formula: "SUM('S1:S3'!A1)" } } }] })
    expect(r.sheets.map((s) => s.cells.C5?.value ?? s.cells.A1.value)).toEqual([{ error: 'VALUE' }, { error: 'VALUE' }, 77, 12])
  })

  it('a session recomputes an intersection when a cell inside it changes, and a structure operation moves its references', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { ...cells, E1: { formula: 'SUM(A1:C3 B2:D4)' } } }] }
    const session = createSession(book)
    expect(session.setCells([{ sheet: 'S', cell: 'B2', input: { value: 10 } }]).changes.find((c) => c.cell === 'E1')?.value).toBe(17)
    const moved = editStructure(book, { op: 'insertRows', sheet: 'S', row: 2 })
    expect(moved.ok && moved.book.sheets[0].cells.E1.formula).toBe('SUM(A1:C4 B3:D5)')
  })
})

describe('what the range operator reads: every cell between its ends', () => {
  const both = (book: BookInput) => {
    const [a, b] = [compute(book), compute(book, { execution: 'eval' })]
    expect(a.hash).toBe(b.hash)
    return a
  }
  const middle: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1 }, B1: { formula: 'E1*10' }, C1: { value: 1 }, E1: { value: 5 }, D1: { formula: 'SUM(A1:A1:C1)' }, D2: { formula: 'SUM(A1:C1)' } } }] }

  it('a formula between the ends is computed before the range is read', () => {
    const c = both(middle).sheets[0].cells
    expect([c.D1.value, c.D2.value]).toEqual([52, 52])
  })

  it('an edit that a formula between the ends depends on reaches the reader in a session, as a full computation gives', () => {
    const session = createSession(middle)
    const { changes } = session.setCells([{ sheet: 'S', cell: 'E1', input: { value: 100 } }])
    expect(changes.find((x) => x.cell === 'D1')?.value).toBe(1002)
    expect(compute(session.getBook()).hash).toBe(session.getResult().hash)
  })

  it('a range around its own cell is a cycle', () => {
    const own = both({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, C1: { value: 7 }, B1: { formula: 'SUM(A1:A1:C1)' } } }] })
    expect([own.sheets[0].cells.B1.value, own.diagnostics]).toEqual([0, [expect.objectContaining({ cell: 'B1', error: 'CYCLE' })]])
    const sheets = both({ sheets: [{ name: 'S1', cells: {} }, { name: 'S3', cells: { A1: { value: 7 }, B1: { formula: 'SUM(S1:S3!A1)' } } }] })
    expect([sheets.sheets[1].cells.B1.value, sheets.diagnostics]).toEqual([0, [expect.objectContaining({ sheet: 'S3', cell: 'B1', error: 'CYCLE' })]])
  })

  it('ends from INDEX: the cells between the references it can answer, not those it only reads', () => {
    const cells = { A1: { value: 1 }, D1: { formula: 'E9*2' }, E9: { value: 4 }, G1: { value: 3 }, H5: { formula: 'SUM(INDEX(A1:C1,1,1):INDEX(E1:G1,1,3))' }, M5: { formula: 'SUM(A1:INDEX(A1:A10,MATCH(5,Z1:Z9,0)))' }, Z3: { value: 5 } }
    const c = both({ sheets: [{ name: 'S', cells }] }).sheets[0].cells
    expect([c.H5.value, c.M5.value]).toEqual([12, 1])
    const session = createSession({ sheets: [{ name: 'S', cells }] })
    expect(session.setCells([{ sheet: 'S', cell: 'E9', input: { value: 10 } }]).changes.find((x) => x.cell === 'H5')?.value).toBe(24)
  })

  it('a spill reference as an end reads as far as the array spills, and follows it when it grows', () => {
    const cells = { A1: { value: 1 }, K1: { value: 3 }, B5: { formula: 'SEQUENCE(1,K1)' }, C1: { formula: 'E9*2' }, E2: { formula: 'E9*100' }, E9: { value: 4 }, J9: { formula: 'SUM(A1:B5#)' } }
    expect(both({ sheets: [{ name: 'S', cells }] }).sheets[0].cells.J9.value).toBe(15)
    const session = createSession({ sheets: [{ name: 'S', cells }] })
    expect(session.setCells([{ sheet: 'S', cell: 'K1', input: { value: 4 } }]).changes.find((x) => x.cell === 'J9')?.value).toBe(419)
    expect(session.setCells([{ sheet: 'S', cell: 'E9', input: { value: 1 } }]).changes.find((x) => x.cell === 'J9')?.value).toBe(113)
    expect(compute(session.getBook()).hash).toBe(session.getResult().hash)
  })

  it('a formula on the middle sheet of a reference across sheets, and in the second area of a union, comes first too', () => {
    const book: BookInput = {
      sheets: [
        { name: 'Jan', cells: { A1: { value: 1 } } },
        { name: 'Feb', cells: { A1: { formula: 'Z!A1*10' } } },
        { name: 'Mar', cells: { A1: { value: 3 } } },
        { name: 'Z', cells: { A1: { value: 5 }, C1: { value: 1 }, D2: { formula: 'A1*10' }, B1: { formula: 'SUM(Jan:Mar!A1)' }, B2: { formula: 'SUBTOTAL(9,(C1:C2,D1:D2))' }, B3: { formula: 'AGGREGATE(9,4,(C1:C2,D1:D2))' }, B4: { formula: 'SUM(FREQUENCY((C1:C2,D1:D2),10))' } } },
      ],
    }
    const z = both(book).sheets[3].cells
    expect([z.B1.value, z.B2.value, z.B3.value, z.B4.value]).toEqual([54, 51, 51, 2])
    const session = createSession(book)
    session.setCells([{ sheet: 'Z', cell: 'A1', input: { value: 7 } }])
    const after = session.getResult().sheets[3].cells
    expect([after.B1.value, after.B2.value, after.B3.value]).toEqual([74, 71, 71])
  })
})
