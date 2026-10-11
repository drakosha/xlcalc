/**
 * The old names of the API, deprecated until hosts move to the new ones: each is the same function as its new name and
 * answers what it answers, from the root, the engine object, a calculator and the three kinds of session.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { createInlineCalc, createSession, editStructure, engine, openSession, restructure, type AsyncRestructureResult, type AsyncSession, type BookInput, type Calc, type RestructureResult, type SessionRestructureResult, type StructureOp } from '../src/index.js'
import { CLOCK } from './async-books.js'
import { closeCalcs, keptPool, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const book = (): BookInput => ({
  sheets: [{
    name: 'S',
    cells: { A1: { value: 2 }, A2: { formula: 'A1*10' }, B1: { value: 'x' } },
    validations: [{ range: 'B1:B3', type: 'list', formula1: '"x,y"' }],
  }],
})
const INSERT: StructureOp = { op: 'insertRows', sheet: 'S', row: 1 }

/** An answer to compare: a result without the time it took. */
const timeless = (x: unknown) => (x && typeof x === 'object' && 'ms' in x ? { ...x, ms: 0 } : x)

/** The session's methods by their old names, with the new name each and its arguments. */
const SESSION: [old: string, name: string, args: unknown[]][] = [
  ['book', 'getBook', []],
  ['layout', 'getLayout', []],
  ['extent', 'getUsedRange', ['S']],
  ['cells', 'getCells', ['S', 'A1:B2']],
  ['validation', 'getValidation', ['S', 'B1']],
  ['result', 'getResult', []],
  ['hash', 'getHash', []],
]

describe('old names of the root and the engine object', () => {
  it('openSession is createSession, restructure is editStructure', () => {
    expect(openSession).toBe(createSession)
    expect(restructure).toBe(editStructure)
    const r: RestructureResult = restructure(book(), INSERT)
    const n = editStructure(book(), INSERT)
    expect(r.ok && n.ok && [r.book, r.mapCell({ sheet: 'S', cell: 'A1' })]).toEqual(n.ok && [n.book, n.mapCell({ sheet: 'S', cell: 'A1' })])
  })

  it('the engine object has both names of each', () => {
    expect(engine.createSession).toBe(createSession)
    expect(engine.openSession).toBe(createSession)
    expect(engine.editStructure).toBe(editStructure)
    expect(engine.restructure).toBe(editStructure)
  })
})

describe('old names of a calculator', () => {
  it.each([['the same thread', () => createInlineCalc()], ['a thread', () => keptPool()]] as const)('%s: openSession is createSession', async (_, make) => {
    const calc: Calc = await make()
    expect(calc.openSession).toBe(calc.createSession)
    const s = await calc.openSession(book(), { clock: CLOCK })
    expect((await s.getResult()).hash).toBe(createSession(book(), { clock: CLOCK }).getHash())
    await s.close()
    if (_ === 'the same thread') await calc.close()
  })
})

describe.each(KINDS)('old names of %s', (kind) => {
  it('each method by its old name is the method of its new name and answers the same', async () => {
    const s = await (await opener(kind))(book(), { clock: CLOCK })
    const methods = s as unknown as Record<string, (...a: unknown[]) => Promise<unknown>>
    for (const [old, name, args] of SESSION) {
      if (kind !== 'a synchronous session') expect(methods[old], old).toBe(methods[name])
      expect(timeless(await methods[old](...args)), old).toEqual(timeless(await methods[name](...args)))
    }
    await s.close()
  })

  it('restructure is editStructure, and undo names the call by its new name', async () => {
    const s: AsyncSession = await (await opener(kind))(book(), { clock: CLOCK })
    if (kind !== 'a synchronous session') expect(s.restructure).toBe(s.editStructure)
    const r: AsyncRestructureResult = await s.restructure(INSERT)
    expect(r.ok && r.mapCell({ sheet: 'S', cell: 'A1' })).toEqual({ sheet: 'S', cell: 'A2' })
    expect((await s.getBook()).sheets[0].cells.A3).toEqual({ formula: 'A2*10' })
    expect((await s.undo())?.call).toBe('editStructure')
    await s.close()
  })
})

describe('old names of a synchronous session', () => {
  it('are the functions of the new names', () => {
    const s = createSession(book(), { clock: CLOCK })
    const methods = s as unknown as Record<string, unknown>
    for (const [old, name] of [...SESSION, ['restructure', 'editStructure']]) expect(methods[old], old).toBe(methods[name])
    const r: SessionRestructureResult = s.restructure(INSERT)
    expect(r).toMatchObject({ ok: true, moved: [{ kind: 'lines' }] })
    expect(r).not.toHaveProperty('book')
  })
})
