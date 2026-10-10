/**
 * A formula's leading `=` is dropped as input is taken: in cells, names, calculated columns of tables, conditional
 * formats and rules of data validation, in `compute`, `restructure`, `validateNames`, `writeXlsx` and a session's edits
 * (in a synchronous session, in the same thread and in a thread), in both execution modes. The book with `=` is the book
 * without it, `inputHash` too; a second `=` stays an error of the formula; the caller's input is not changed.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, createInlineCalc, openSession, restructure, validateNames, type AsyncSession, type BookInput, type Calc, type Execution, type Session, type SessionOptions } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { CLOCK } from './async-books.js'
import { pool } from './fixtures/pool.js'

/** The book with `=` before each of its formulas, or without (`eq` empty). */
const book = (eq: string): BookInput => ({
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 'Qty' }, A2: { value: 2 }, A3: { value: 5 },
        B1: { value: 'Twice' }, B2: { formula: `${eq}[@Qty]*2` }, B3: { formula: `${eq}[@Qty]*2` },
        C1: { formula: `${eq}SUM(A2:A3)` },
        C2: { formula: `${eq}rate*C1` },
        D1: { value: 'Open' },
      },
      tables: [{ name: 'T', range: 'A1:B3', columns: [{ name: 'Qty' }, { name: 'Twice', formula: `${eq}[@Qty]*2` }] }],
      conditionalFormats: [
        { type: 'cellIs', ranges: ['A2:A3'], operator: 'greaterThan', formulas: [`${eq}3`], style: { font: { bold: true } } },
        { type: 'expression', ranges: ['C1'], formula: `${eq}C1>5`, style: { fill: { color: '#FFC7CE' } } },
        { type: 'dataBar', ranges: ['A2:A3'], min: { type: 'num', value: `${eq}0` }, max: { type: 'max' }, color: '#638EC6' },
      ],
      validations: [{ range: 'D1:D9', type: 'list', formula1: `${eq}"Open,Closed"` }, { range: 'E1', type: 'whole', operator: 'between', formula1: `${eq}1`, formula2: `${eq}$A$3` }],
    },
  ],
  names: [{ name: 'rate', ref: `${eq}0.5` }],
})

describe.each(['closures', 'eval'] as Execution[])('a leading = in the input, %s', (execution) => {
  it('is dropped: the book computes as the one without it, with the same input hash and the same values', () => {
    const [plain, eq] = [compute(book(''), { clock: CLOCK, execution }), compute(book('='), { clock: CLOCK, execution })]
    expect(eq.issues).toEqual([])
    expect(eq.inputHash).toBe(plain.inputHash)
    expect(eq.hash).toBe(plain.hash)
    expect([eq.sheets[0].cells.C1.value, eq.sheets[0].cells.C2.value, eq.sheets[0].cells.B3.value]).toEqual([7, 3.5, 10])
    expect(eq.sheets[0].cells.C1.formula).toBe('SUM(A2:A3)')
  })

  it("leaves the caller's book as it is, and a book without = as it was", () => {
    const b = book('=')
    compute(b, { clock: CLOCK, execution })
    expect(b.sheets[0].cells.C1.formula).toBe('=SUM(A2:A3)')
    expect(b.names![0].ref).toBe('=0.5')
    const s = openSession(book(''), { clock: CLOCK, execution })
    expect(s.book().sheets[0].cells.C1).toEqual({ formula: 'SUM(A2:A3)' })
    s.close()
  })

  it('a second = stays: an error of the formula', () => {
    const r = compute({ sheets: [{ name: 'S', cells: { A1: { formula: '==1' }, A2: { formula: '=' } } }] }, { clock: CLOCK, execution })
    expect([r.sheets[0].cells.A1.value, r.sheets[0].cells.A2.value]).toEqual([{ error: 'ERROR' }, { error: 'ERROR' }])
    expect(r.issues.map((i) => `${i.cell} ${i.check}`)).toEqual(['A1 syntax', 'A2 empty'])
    expect(r.sheets[0].cells.A1.formula).toBe('=1')
  })

  it('restructure, validateNames and writeXlsx take it so too', async () => {
    const moved = restructure(book('='), { op: 'insertRows', sheet: 'S', row: 1 })
    if (!moved.ok) throw new Error(moved.message)
    expect(moved.book.sheets[0].cells.C2.formula).toBe('SUM(A3:A4)')
    const added = restructure(book(''), { op: 'addTable', sheet: 'S', table: { name: 'U', range: 'F1:G2', columns: [{ name: 'a' }, { name: 'b', formula: '=[@a]+1' }] } })
    if (!added.ok) throw new Error(added.message)
    expect(added.book.sheets[0].tables![1].columns[1].formula).toBe('[@a]+1')
    expect(validateNames([{ name: 'first', ref: "='S'!$A$1" }], { sheets: ['S'] })).toEqual([])
    const b = book('=')
    const read = await readXlsx((await writeXlsx({ book: b, result: compute(b, { clock: CLOCK, execution }) })).data)
    expect(read.issues).toEqual([])
    expect(read.book.sheets[0].cells.C1).toEqual({ formula: 'SUM(A2:A3)' })
    expect(read.book.names).toEqual([{ name: 'rate', ref: '0.5' }])
  })
})

/** A synchronous session behind the asynchronous API, so that one test covers the three ways. */
function wrapped(s: Session): AsyncSession {
  const out: Record<string, unknown> = {}
  for (const k of ['setCells', 'setNames', 'setLayout', 'restructure', 'validateEntry', 'book', 'result', 'close'] as const)
    out[k] = async (...a: unknown[]) => (s[k] as (...a: unknown[]) => unknown)(...a)
  return out as unknown as AsyncSession
}

const KINDS = ['a synchronous session', 'the same thread', 'a thread'] as const
type Kind = (typeof KINDS)[number]

const calcs: Calc[] = []
afterAll(() => Promise.all(calcs.map((c) => c.close())))

async function opener(kind: Kind): Promise<(book: BookInput, opts: SessionOptions) => Promise<AsyncSession>> {
  if (kind === 'a synchronous session') return async (b, opts) => wrapped(openSession(b, opts))
  const calc = kind === 'the same thread' ? createInlineCalc({ sliceMs: 2 }) : await pool({ threads: 1 })
  calcs.push(calc)
  return (b, opts) => calc.openSession(b, opts)
}

describe.each(KINDS)("a session's edits, %s", (kind) => {
  it.each(['closures', 'eval'] as Execution[])('drop a leading = as the book does, %s', async (execution) => {
    const [open, eq] = [await opener(kind), await opener(kind)]
    const [a, b] = await Promise.all([open(book(''), { clock: CLOCK, execution }), eq(book('='), { clock: CLOCK, execution })])
    expect((await b.result()).inputHash).toBe((await a.result()).inputHash)
    expect(await b.setCells([{ sheet: 'S', cell: 'C3', input: { formula: '=C1*10' } }])).toEqual(await a.setCells([{ sheet: 'S', cell: 'C3', input: { formula: 'C1*10' } }]))
    expect(await b.setNames([{ name: 'rate', ref: '=2' }])).toEqual(await a.setNames([{ name: 'rate', ref: '2' }]))
    const cf = (e: string) => [{ sheet: 'S', conditionalFormats: [{ type: 'expression' as const, ranges: ['C1'], formula: `${e}C1>1`, style: { fill: { color: '#FFC7CE' } } }] }]
    await Promise.all([b.setLayout(cf('=')), a.setLayout(cf(''))])
    const op = (e: string) => ({ op: 'addTable' as const, sheet: 'S', table: { name: 'U', range: 'F1:G2', columns: [{ name: 'a' }, { name: 'b', formula: `${e}[@a]+1` }] } })
    await Promise.all([b.restructure(op('=')), a.restructure(op(''))])
    expect(await b.validateEntry('S', 'E1', { formula: '=A3' })).toEqual({ ok: true })
    expect(await b.validateEntry('S', 'E1', { formula: '=A3+1' })).toEqual(await a.validateEntry('S', 'E1', { formula: 'A3+1' }))
    const [x, y] = [await a.book(), await b.book()]
    expect(y).toEqual(x)
    expect(y.sheets[0].cells.C3).toEqual({ formula: 'C1*10' })
    expect((await b.result()).hash).toBe((await a.result()).hash)
    expect((await b.setCells([{ sheet: 'S', cell: 'C3', input: { formula: '==1' } }]))[0].value).toEqual({ error: 'ERROR' })
    await Promise.all([a.close(), b.close()])
  })
})
