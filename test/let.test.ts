/**
 * LET beyond Excel's answers (`test/excel/cases/let*.json`, `test/excel/spill/let.json`): the scope the parser
 * resolves, the canonical text, the checks and the calls Excel does not load, values computed once and on demand,
 * references kept, the graph through the values, spills, sessions, ru-RU text and structure operations.
 */
import { describe, expect, it } from 'vitest'
import { REFUSED as HELD_REFUSED } from '../bench/coverage/sets/let-held.js'
import { REFUSED } from '../bench/coverage/sets/let.js'
import { compute, openSession, restructure, ruRU, toInvariant, toLocal, validateFormula, type BookInput } from '../src/index.js'
import { canonical, parse, type Ast } from '../src/syntax.js'
import { engineFormula } from './excel-cases.js'

/** A sheet S: `=…` a formula, else a value. */
const sheet = (cells: Record<string, string | number>): BookInput => ({
  sheets: [{ name: 'S', cells: Object.fromEntries(Object.entries(cells).map(([a, v]) => [a, typeof v === 'string' && v.startsWith('=') ? { formula: v.slice(1) } : { value: v }])) }],
})

/** The uses of names in a formula: LET's by slot, the book's by name. */
function uses(ast: Ast): string[] {
  const out: string[] = []
  const go = (n: Ast) => {
    if (n.k === 'local') out.push(`${n.name}@${n.slot}`)
    else if (n.k === 'name') out.push(n.name)
    else if (n.k === 'fn') n.args.forEach(go)
    else if (n.k === 'bin') [n.a, n.b].forEach(go)
    else if (n.k === 'un' || n.k === 'pct' || n.k === 'paren') go(n.a)
    else if (n.k === 'let') [...n.binds.map((b) => b.value), ...(n.body ? [n.body] : [])].forEach(go)
  }
  go(ast)
  return out
}

describe('LET', () => {
  it('a name is seen from the end of its value to the end of its LET; an inner LET shadows it', () => {
    expect(uses(parse('LET(x,1,y,x+1,LET(x,y,x*2)+x)'))).toEqual(['x@0', 'y@1', 'x@2', 'x@0'])
    // in its own value and past its LET the word is a defined name
    expect(uses(parse('LET(x,x+1,x)+x'))).toEqual(['x', 'x@0', 'x'])
  })

  it('canonical text: LET upper-case, every use spelled as its name is declared, as Excel shows it', () => {
    expect(canonical(parse('let(Abc,1,ABC+abc)'))).toBe('LET(Abc,1,Abc+Abc)')
    expect(canonical(parse('LET( x , , x )'))).toBe('LET(x,,x)')
    expect(canonical(parse('LET(x,1,y,2)'))).toBe('LET(x,1,y,2)')
  })

  it('checks: LET among the functions, its names not the book’s', () => {
    expect(validateFormula('LET(x,A1,SUM(x,nm))', { names: ['nm'] })).toMatchObject({ ok: true, functions: ['LET', 'SUM'], names: ['nm'] })
    expect(validateFormula('LET(x,1,x)+x', { names: [] })).toMatchObject({ ok: false, code: 'unknown_name', message: 'unknown name x' })
  })

  it('Excel does not load a name twice, a reference, number or text as a name, no name, over 254 arguments', () => {
    for (const f of [...REFUSED, ...HELD_REFUSED]) expect(validateFormula(engineFormula(f)), f).toMatchObject({ ok: false, code: 'syntax' })
    expect(validateFormula('LET(x,1,X,2,x)')).toMatchObject({ message: 'LET names X twice' })
    expect(validateFormula('LET(1,2,3)')).toMatchObject({ message: 'LET takes a name as its argument 1' })
    expect(validateFormula('LET(TRUE,1,2)').ok).toBe(false)
  })

  it('a value is computed once, on its first use; an unused one does not count', () => {
    const r = compute(sheet({ A1: '=LET(x,RAND(),x-x)', A2: '=LET(x,1/0,1)', A3: '=LET(x,NA(),y,5,IF(FALSE,x,y))' }), { clock: { now: '2026-10-07T00:00:00Z', seed: 1 } })
    expect(['A1', 'A2', 'A3'].map((a) => r.sheets[0].cells[a].value)).toEqual([0, 1, 5])
    expect(r.volatile.map((v) => v.cell)).toEqual(['A1'])
  })

  it('a reference stays a reference: SUM skips text, SUMIF takes it, LET answers it', () => {
    const r = compute(sheet({ A1: '2', A2: 3, B1: '=LET(r,A1:A2,SUM(r))', B2: '=SUMIF(LET(r,A1:A2,r),">0")', B3: '=LET(r,A1,r&"")', B4: '=ROWS(LET(r,A1:A2,r))' }))
    expect(['B1', 'B2', 'B3', 'B4'].map((a) => r.sheets[0].cells[a].value)).toEqual([3, 3, '2', 2])
  })

  it('the references in its values are the formula’s: an edit recomputes it', () => {
    const s = openSession(sheet({ A1: 1, B1: '=LET(x,A1*10,x+1)' }))
    expect(s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 2 } }]).changes.map((c) => [c.cell, c.value])).toContainEqual(['B1', 21])
    s.close()
  })

  it('an array result spills', () => {
    const r = compute(sheet({ A1: '=LET(s,SEQUENCE(3),s*2)' }))
    expect(['A1', 'A2', 'A3'].map((a) => r.sheets[0].cells[a].value)).toEqual([2, 4, 6])
  })

  it('ru-RU text: LET keeps its name, its names stay as they are', () => {
    expect(toLocal('LET(x,1.5,SUM(x,1))', ruRU)).toMatchObject({ ok: true, formula: 'LET(x;1,5;СУММ(x;1))' })
    expect(toInvariant('=LET(сумма;2;сумма*2)', ruRU)).toMatchObject({ ok: true, formula: 'LET(сумма,2,сумма*2)' })
  })

  it('structure operations move the references in its values, not its names', () => {
    const book = sheet({ A1: 1, B1: '=LET(A,A1,b,A+1,b*$A$1)' })
    const r = restructure(book, { op: 'insertRows', sheet: 'S', row: 1 })
    expect(r.ok && r.book.sheets[0].cells.B2.formula).toBe('LET(A,A2,b,A+1,b*$A$2)')
    const renamed = restructure({ sheets: [...book.sheets, { name: 'T', cells: { A1: { formula: 'LET(x,S!A1,x)' } } }] }, { op: 'renameSheet', sheet: 'S', name: 'R' })
    expect(renamed.ok && renamed.book.sheets[1].cells.A1.formula).toBe("LET(x,'R'!A1,x)")
  })
})
