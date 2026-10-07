/**
 * Names of any formula beyond Excel's answers (`test/excel/spill/names*.json`): the issues of names the engine does
 * not compute, cycles through names, LET's names of a name apart from the formula's, references wrapping round the
 * sheet, names qualified by a sheet in checks, a session's `setNames`, and both execution modes.
 */
import { describe, expect, it } from 'vitest'
import { compute, openSession, validateFormula, type BookInput, type NameInput } from '../src/index.js'

/** A sheet S with values A1:A3 = 1, 2, 3 and the formulas given, and the names. */
function book(cells: Record<string, string>, names: NameInput[], more: BookInput['sheets'] = []): BookInput {
  return {
    sheets: [{ name: 'S', cells: { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 }, ...Object.fromEntries(Object.entries(cells).map(([a, f]) => [a, { formula: f }])) } }, ...more],
    names,
  }
}

const values = (b: BookInput, sheet = 'S') => Object.fromEntries(Object.entries(compute(b, { display: false }).sheets.find((s) => s.name === sheet)!.cells).map(([a, c]) => [a, c.value]))

describe('names of any formula', () => {
  it('a formula that does not parse, or a reference without its sheet, leaves the name out with its issue', () => {
    const r = compute(book({ B1: 'bad', B2: 'nosheet' }, [{ name: 'bad', ref: "'S'!A1+" }, { name: 'nosheet', ref: 'A1*2' }]), { display: false })
    expect(r.nameIssues.map((i) => `${i.name}: ${i.reason}`)).toEqual([expect.stringMatching(/^bad: bad formula/), 'nosheet: a reference in a name needs its sheet'])
    expect(r.issues.map((i) => `${i.cell} ${i.check}`)).toEqual(['B1 unknown_name', 'B2 unknown_name'])
  })

  it('a formula over an unknown sheet or function stays a name giving #REF! or #NAME?, with its issue', () => {
    const r = compute(book({ B1: 'gone+1', B2: 'udf' }, [{ name: 'gone', ref: "'X'!A1" }, { name: 'udf', ref: 'FOO(1)' }]), { display: false })
    expect(r.sheets[0].cells.B1.value).toEqual({ error: 'REF' })
    expect(r.sheets[0].cells.B2.value).toEqual({ error: 'NAME' })
    expect(r.nameIssues.map((i) => i.name)).toEqual(['gone', 'udf'])
  })

  it('a name that uses itself, through other names or not, is #CYCLE! where it is used', () => {
    const v = values(book({ B1: 'a', B2: 'self+1', B3: 'SUM(A1:A3)' }, [{ name: 'a', ref: 'b+1' }, { name: 'b', ref: 'a+1' }, { name: 'self', ref: 'self*2' }]))
    expect([v.B1, v.B2, v.B3]).toEqual([{ error: 'CYCLE' }, { error: 'CYCLE' }, 6])
  })

  it("a name's LET names are its own, apart from a LET of the formula using it", () => {
    const v = values(book({ B1: 'LET(x,10,y,20,x+y+lt)', B2: 'LET(a,100,a+lt)' }, [{ name: 'lt', ref: "LET(a,'S'!$A$2,b,a*3,a+b)" }]))
    expect([v.B1, v.B2]).toEqual([38, 108])
  })

  it('a reference without $ moves by the cell using the name and wraps round the edges of the sheet', () => {
    const names = [{ name: 'above', ref: "'S'!A1048576" }, { name: 'left', ref: "'S'!XFD1" }]
    const v = values(book({ A4: 'above*10', B3: 'left+1', C1: 'above' }, names))
    // C1 reads C1048576, blank
    expect([v.A4, v.B3, v.C1]).toEqual([30, 4, 0])
  })

  it('a name qualified by a sheet is that sheet\'s local name, else the global one; the check knows them', () => {
    const names = [{ name: 'k', ref: '1' }, { name: 'k', ref: '5', scope: 'T' }]
    const b = book({ B1: "'T'!k*10", B2: 'k', B3: "'S'!k" }, names, [{ name: 'T', cells: { A1: { formula: 'k' } } }])
    expect(values(b)).toMatchObject({ B1: 50, B2: 1, B3: 1 })
    expect(values(b, 'T').A1).toBe(5)
    expect(validateFormula("'T'!k+'U'!k", { sheets: ['S', 'T'], names: ['k'] })).toMatchObject({ ok: false, code: 'unknown_sheet' })
    expect(validateFormula("'T'!k", { sheets: ['S', 'T'], names: ['k'] })).toMatchObject({ ok: true, canonical: "'T'!k", names: ['k'] })
  })

  it('a session takes new names of formulas and returns the cells they change', () => {
    const s = openSession(book({ B1: 'rate*100' }, [{ name: 'rate', ref: '0.2' }]))
    expect(s.result().sheets[0].cells.B1.value).toBe(20)
    expect(s.setNames([{ name: 'rate', ref: "'S'!$A$3/10" }]).map((c) => [c.cell, c.value])).toEqual([['B1', 30]])
  })

  it('both execution modes give the same values over names of formulas', () => {
    const b = book({ B1: 'u+SUM(un)+INDEX(un,1,1,2)+rel' }, [{ name: 'u', ref: "SUM('S'!$A$1:$A$3)*2" }, { name: 'un', ref: "'S'!$A$1,'S'!$A$3" }, { name: 'rel', ref: "'S'!XFD1" }])
    const [x, y] = [compute(b, { execution: 'eval', display: false }), compute(b, { execution: 'closures', display: false })]
    expect(x.sheets).toEqual(y.sheets)
    expect(x.sheets[0].cells.B1.value).toBe(12 + 4 + 3 + 1)
  })
})
