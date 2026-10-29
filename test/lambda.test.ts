/**
 * LAMBDA beyond Excel's answers (`test/excel/cases/lambda*.json`, `groupby*.json`, `test/excel/workbook/lambda-names.json`):
 * the parser's parameters and calls, the checks, names that are LAMBDAs, functions as values, both execution modes,
 * every kind of session with its edits, undo and structure operations, the depth of recursion and a deadline, xlsx and
 * writing over a file Excel saved, ru-RU text, `explain` and the tips.
 */
import { readFileSync } from 'node:fs'
import { afterAll, describe, expect, it } from 'vitest'
import { compute, createInlineCalc, createSession, functionSignature, toInvariant, toLocal, validateFormula, validateLocalFormula, type BookInput, type CellValue } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { canonical, parse } from '../src/syntax.js'
import { ruRU } from './locales.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'
import { ruRUFunctionText } from '../src/function-text/ru-RU.js'

const values = (book: BookInput, cells: string[], execution: 'closures' | 'eval' = 'closures'): CellValue[] => {
  const out = compute(book, { execution }).sheets[0].cells
  return cells.map((a) => out[a]?.value ?? null)
}

/** A sheet S of values and `=` formulas, with names. */
const book = (cells: Record<string, string | number>, names: BookInput['names'] = []): BookInput => ({
  sheets: [{ name: 'S', cells: Object.fromEntries(Object.entries(cells).map(([a, v]) => [a, typeof v === 'string' && v.startsWith('=') ? { formula: v.slice(1) } : { value: v }])) }],
  names,
})

const NAMES: BookInput['names'] = [
  { name: 'Twice', ref: 'LAMBDA(x,x*2)' },
  { name: 'Cnt', ref: 'LAMBDA(n,IF(n<=0,0,1+Cnt(n-1)))' },
  { name: 'EvenQ', ref: 'LAMBDA(n,IF(n=0,TRUE,OddQ(n-1)))' },
  { name: 'OddQ', ref: 'LAMBDA(n,IF(n=0,FALSE,EvenQ(n-1)))' },
  { name: 'PlusA1', ref: "LAMBDA(x,x+'S'!$A$1)" },
]

describe('LAMBDA: parsing and checks', () => {
  it('parameters are seen in the calculation, calls follow a LAMBDA, a call or a LET name, and print as typed', () => {
    for (const f of ['LAMBDA(x,y,x+y)(1,2)', 'LAMBDA(x,LAMBDA(y,x+y))(1)(2)', 'LET(f,LAMBDA(x,x+1),f(2))', '(LAMBDA(x,x))(3)', 'LAMBDA(a1,a1*2)(4)']) expect(canonical(parse(f))).toBe(f)
    const ast = parse('LAMBDA(x,x+y)')
    expect(ast.k === 'lambda' && ast.body.k === 'bin' && [ast.body.a.k, ast.body.b.k]).toEqual(['local', 'name'])
  })

  it('refuses a parameter twice, one with a period, a LAMBDA without its calculation, and more than 253 parameters', () => {
    expect(validateFormula('LAMBDA(x,X,1)')).toMatchObject({ ok: false, code: 'syntax', message: 'LAMBDA names the parameter X twice' })
    expect(validateFormula('LAMBDA(a.b,1)')).toMatchObject({ ok: false, code: 'syntax' })
    expect(validateFormula('LAMBDA(x,)')).toMatchObject({ ok: false, message: 'LAMBDA needs a calculation as its last argument' })
    expect(validateFormula(`LAMBDA(${Array.from({ length: 254 }, (_, i) => `p_${i}`).join(',')},1)`)).toMatchObject({ ok: false, message: 'LAMBDA takes at most 253 parameters' })
    expect(validateFormula(`LAMBDA(${Array.from({ length: 253 }, (_, i) => `p_${i}`).join(',')},1)`).ok).toBe(true)
  })

  it('a call of a name the book defines is the name\'s, spelled as defined; a function\'s name alone is the function', () => {
    expect(validateFormula('twice(2)', { names: ['Twice'] })).toMatchObject({ ok: true, names: ['TWICE'], functions: [] })
    expect(validateFormula('twice(2)')).toMatchObject({ ok: false, code: 'unknown_function' })
    expect(validateFormula('BYROW(A1:C3,sum)')).toMatchObject({ ok: true, canonical: 'BYROW(A1:C3,sum)', functions: ['BYROW', 'SUM'], names: [] })
    // the text keeps the word, as the book may define it later: then it is the name
    expect(values(book({ A1: '=Rate+1' }), ['A1'])).toEqual([{ error: 'VALUE' }])
    const named = compute(book({ A1: '=Rate+1' }, [{ name: 'Rate', ref: '2' }])).sheets[0].cells.A1
    expect([named.formula, named.value]).toEqual(['Rate+1', 3])
    // a LET name of a function's word calls the function
    expect(values(book({ A1: '=LET(sum,5,SUM(sum,1))' }), ['A1'])).toEqual([6])
    const out = compute(book({ A1: '=twice(3)' }, NAMES)).sheets[0].cells.A1
    expect([out.formula, out.value]).toEqual(['Twice(3)', 6])
  })
})

describe('LAMBDA: values in both execution modes', () => {
  const cases: [string, CellValue][] = [
    ['=Twice(3)', 6],
    ['=Cnt(150)', 150],
    ['=EvenQ(31)', false],
    // a name of a function's word calls the function, as Excel reads it
    ['=LET(x,1,ISODD(3))', true],
    ['=PlusA1(1)', 3],
    ['=LET(a,10,LAMBDA(x,x+a))(5)', 15],
    ['=LAMBDA(x,y,IF(ISOMITTED(y),"no",y))(1,)', 'no'],
    ['=LAMBDA(x,x)', { error: 'CALC' }],
    ['=TYPE(SUM)', 128],
    ['=SUM(MAP(A1:A3,LAMBDA(x,x*2)))', 12],
    ['=REDUCE(,A1:A3,LAMBDA(a,b,a&b))', '231'],
    ['=SUM(BYROW(A1:B3,SUM))', 6],
    ['=ROWS(GROUPBY(B1:B3,A1:A3,SUM))', 3],
    // a LET inside a LAMBDA calling itself keeps each call's own values
    ['=LET(f,LAMBDA(g,n,LET(m,n-1,IF(n<=0,0,g(g,m)+m))),f(f,4))', 6],
  ]
  for (const execution of ['closures', 'eval'] as const)
    it(execution, () => {
      const b = book({ A1: 2, A2: 3, A3: 1, B1: 'p', B2: 'q', B3: 'p', ...Object.fromEntries(cases.map(([f], i) => [`D${i + 1}`, f])) }, NAMES)
      expect(values(b, cases.map((_, i) => `D${i + 1}`), execution)).toEqual(cases.map(([, v]) => v))
    })

  it('a formula with a LAMBDA copied down computes at each cell: no shape is shared', () => {
    const b = book({ A1: 1, A2: 2, A3: 3, B1: '=LAMBDA(x,x*A1)(10)', B2: '=LAMBDA(x,x*A2)(10)', B3: '=LAMBDA(x,x*A3)(10)' })
    expect(values(b, ['B1', 'B2', 'B3'], 'eval')).toEqual([10, 20, 30])
  })

  it('a recursion past the budget is #NUM!, and a long one stops at the deadline', async () => {
    expect(values(book({ A1: '=Cnt(100000)' }, NAMES), ['A1'])).toEqual([{ error: 'NUM' }])
    const calc = createInlineCalc()
    const slow = book({ A1: '=SUM(MAKEARRAY(2000,2000,LAMBDA(r,c,r*c)))' })
    await expect(calc.compute(slow, { timeoutMs: 20 })).rejects.toMatchObject({ code: 'soft_timeout' })
    await calc.close()
  })
})

describe('LAMBDA in names: every kind of session', () => {
  afterAll(closeCalcs)
  for (const kind of KINDS)
    it(`${kind}: a cell the calculation reads, the names set again, undone, rows inserted`, async () => {
      const s = await (await opener(kind))(book({ A1: 2, B1: '=PlusA1(1)', B2: '=Twice(B1)' }, NAMES), {})
      const read = async () => Object.values(await s.getCells('S', ['B1', 'B2', 'B3'])).map((c) => c.value ?? null)
      expect(await read()).toEqual([3, 6])
      await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 10 } }])
      expect(await read()).toEqual([11, 22])
      await s.setNames([...NAMES!.filter((n) => n.name !== 'Twice'), { name: 'Twice', ref: 'LAMBDA(x,x*3)' }])
      expect(await read()).toEqual([11, 33])
      await s.undo()
      expect(await read()).toEqual([11, 22])
      await s.editStructure({ op: 'insertRows', sheet: 'S', row: 1, count: 1 })
      const names = (await s.getBook()).names!
      expect(names.find((n) => n.name === 'PlusA1')!.ref).toBe("LAMBDA(x,x+'S'!$A$2)")
      expect(Object.values(await s.getCells('S', ['B2', 'B3'])).map((c) => c.value)).toEqual([11, 22])
      await s.close?.()
    })
})

describe('LAMBDA: files, locales, explain and tips', () => {
  it('writes and reads LAMBDAs in names and cells, a function as a value with its prefix', async () => {
    const b = book({ A1: 2, B1: '=Twice(A1)', B2: '=BYROW(A1:A1,SUM)', B3: '=LAMBDA(x,y,x*y)(3,4)' }, NAMES)
    const w = await writeXlsx({ book: b, result: compute(b) })
    const r = await readXlsx(w.data)
    expect(r.book.names).toEqual(NAMES)
    expect(['B1', 'B2', 'B3'].map((a) => r.book.sheets[0].cells[a].formula)).toEqual(['Twice(A1)', 'BYROW(A1:A1,SUM)', 'LAMBDA(x,y,x*y)(3,4)'])
  })

  it('writing over the file Excel saved keeps its LAMBDAs in names and cells', async () => {
    const file = readFileSync(new URL('./excel/xlsx/roundtrip/lambdas.xlsx', import.meta.url))
    const kept = await openKept(file)
    const s = createSession(kept.book, { cached: kept.cached })
    s.setCells([{ sheet: 'L', cell: 'B1', input: { value: 11 } }])
    const again = await readXlsx((await writeKept(file, s.getSnapshot())).data)
    s.close()
    const own = (names: BookInput['names']) => names?.filter((n) => !n.hidden)
    expect(own(again.book.names)).toEqual(own(kept.book.names))
    expect(again.book.sheets[0].cells.C2.formula).toBe('Twice(B1)+Bump(1)')
    expect(compute(again.book).sheets[0].cells.C2.value).toBe(28)
  })

  it('ru-RU keeps a function as a value in English, as a word alone may be a name (Ставка, СТАВКА being RATE)', () => {
    expect(toLocal('BYROW(A1:C3,SUM)', ruRU)).toEqual({ ok: true, formula: 'BYROW(A1:C3;SUM)' })
    expect(toInvariant('BYROW(A1:C3;SUM)', ruRU)).toEqual({ ok: true, formula: 'BYROW(A1:C3,SUM)' })
    expect(toLocal('MAP(A1:A3,LAMBDA(sum,sum*2))', ruRU)).toEqual({ ok: true, formula: 'MAP(A1:A3;LAMBDA(sum;sum*2))' })
    // a word the book defines stays its name: СТАВКА is RATE in ru-RU
    expect(validateLocalFormula('=BYROW(A1:C3;SUM)+Ставка', ruRU, { names: ['Ставка'] })).toMatchObject({ ok: true, canonical: 'BYROW(A1:C3,SUM)+Ставка' })
    expect(toLocal('LAMBDA(x,y,IF(ISOMITTED(y),x,y))(1,)', ruRU)).toEqual({ ok: true, formula: 'LAMBDA(x;y;ЕСЛИ(ПРОПУЩЕНО(y);x;y))(1;)' })
  })

  it('explain leads a call of a name to the cells its LAMBDA reads', () => {
    const s = createSession(book({ A1: 2, B1: '=PlusA1(1)' }, NAMES))
    const e = s.explain('S', 'B1')
    expect(e.root.children.map((c) => c.ref)).toEqual(["'S'!A1"])
    s.close()
  })

  it('tips and descriptions', () => {
    expect(functionSignature('LAMBDA')!.args.map((a) => a.name)).toEqual(['parameter_or_calculation1', 'parameter_or_calculation2'])
    expect(functionSignature('GROUPBY')).toMatchObject({ category: 'lookup' })
    expect(functionSignature('ГРУПППО', { ...ruRU, ...ruRUFunctionText })).toMatchObject({ fn: 'GROUPBY' })
  })
})
