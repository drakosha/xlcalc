/**
 * Live objects of a synchronous session's sheets and cells (`liveBook`, `src/live.ts`), in both modes of execution:
 * a read after an assignment gives what the session recomputed, an assignment is the session's call (its refusal too,
 * the book unchanged), a batch is one `applyChanges`; a cell's listener is called for the changes that touched it,
 * whoever made them, and not for others; a structure operation wakes the cells of its sheet whose address other cells
 * came to; a sheet follows its rename and ends with its deletion.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { CalcError, createInlineCalc, createSession, InputError, liveBook, type BookInput, type Execution, type CalcLimits, type SessionChange } from '../src/index.js'
import { clock } from './struct-sessions.js'

const book = (): BookInput => ({
  sheets: [
    { name: 'Plan', cells: { A1: { value: 2 }, A2: { formula: 'A1*10' }, A5: { value: 5 }, C1: { value: 'c' } } },
    { name: 'Other', cells: { A1: { formula: 'Plan!A1+1' } } },
  ],
})

/** Every change a listener is called with, by revision. */
function heard() {
  const got: number[] = []
  return { got, listener: (c: SessionChange) => got.push(c.revision) }
}

afterEach(() => vi.unstubAllGlobals())

describe.each(['closures', 'eval'] as const)('live objects, %s', (execution: Execution) => {
  const open = (opts: { display?: boolean; limits?: CalcLimits } = {}) => {
    const session = createSession(book(), { clock, execution, ...opts })
    const live = liveBook(session)
    return { session, live, plan: live.sheet('Plan') }
  }

  it('a read after an assignment gives what the session recomputed, the dependent cells too; each assignment is one step of the undo', () => {
    const { session, plan, live } = open()
    plan.cells.A1.value = 3
    expect([plan.cells.A1.value, plan.cells.A2.value, live.sheet('Other').cells.A1.value]).toEqual([3, 30, 4])
    plan.cells.B1.formula = 'A1*1.5'
    expect([plan.cells.B1.value, plan.cells.B1.formula, plan.cells.B1.text]).toEqual([4.5, 'A1*1.5', '4.5'])
    expect(session.history.undo).toBe(2)
    // a value over a formula takes its place; null clears the content
    plan.cells.A2.value = 'x'
    expect([plan.cells.A2.value, plan.cells.A2.formula]).toEqual(['x', undefined])
    plan.cells.A1.value = null
    expect([plan.cells.A1.value, plan.cells.A1.text, session.getCell('Plan', 'A1'), plan.cells.B1.value]).toEqual([null, undefined, undefined, 0])
    plan.cells.B1.formula = null
    expect(plan.cells.B1.value).toBe(null)
    expect(session.history.undo).toBe(5)
    session.close()
  })

  it('the look by patchCells: the value stays; a value typed over it keeps it, null removes it', () => {
    const { session, plan } = open()
    const a1 = plan.cells.A1
    a1.numberFormat = '0.00'
    a1.style = { font: { bold: true } }
    expect([a1.value, a1.text, a1.numberFormat, a1.style, session.getLayout().styles]).toEqual([2, '2.00', '0.00', 0, [{ font: { bold: true } }]])
    a1.value = 4
    expect([a1.text, a1.numberFormat, a1.style]).toEqual(['4.00', '0.00', 0])
    a1.numberFormat = null
    a1.style = null
    expect([a1.text, a1.numberFormat, a1.style]).toEqual(['4', undefined, undefined])
    session.close()
  })

  it('text typed into a cell, as setCellText answers it', () => {
    const { session, plan } = open()
    expect(plan.cells.B2.type('5%')).toMatchObject({ ok: true, revision: 1 })
    expect([plan.cells.B2.value, plan.cells.B2.text]).toEqual([0.05, '5%'])
    expect(plan.cells.A2.type('=A1+')).toMatchObject({ ok: false })
    expect(session.revision).toBe(1)
    session.close()
  })

  it('without display a cell has no text, as getCell has none', () => {
    const { session, plan } = open({ display: false })
    expect([plan.cells.A1.value, plan.cells.A1.text]).toEqual([2, undefined])
    session.close()
  })

  it("an assignment the session refuses throws its call's error and leaves the book as it was", () => {
    const { session, live, plan } = open({ limits: { writes: 1 } })
    session.setCells([{ sheet: 'Plan', cell: 'D1', input: { formula: '{1;2}', array: 'D1:D2' } }])
    const before = [session.getHash(), session.revision]
    const refusals: [() => void, unknown][] = [
      // part of an array of the old kind
      [() => (plan.cells.D2.value = 1), InputError],
      [() => (plan.cells.B9.value = { bad: 1 } as never), InputError],
      // over the limit of writes: two cells in one batch
      [
        () =>
          live.batch(() => {
            plan.cells.A1.value = 1
            plan.cells.A3.value = 1
          }),
        CalcError,
      ],
    ]
    for (const [assign, error] of refusals) {
      expect(assign).toThrow(error as typeof Error)
      expect([session.getHash(), session.revision]).toEqual(before)
    }
    // the session's own error, word for word
    let direct: unknown
    try {
      session.setCells([{ sheet: 'Plan', cell: 'D2', input: { value: 1 } }])
    } catch (e) {
      direct = e
    }
    expect(() => (plan.cells.D2.value = 1)).toThrow(direct as Error)
    session.close()
  })

  it('a batch is one applyChanges: one recomputation, one revision, one step of the undo, one change; until it ends the cells read as before', () => {
    const { session, live, plan } = open()
    const changes: SessionChange[] = []
    session.subscribe((c) => changes.push(c))
    let inside: unknown[] = []
    live.batch(() => {
      plan.cells.A1.value = 7
      plan.cells.B1.formula = 'A2+1'
      plan.cells.A1.numberFormat = '0.0'
      // a batch inside another is part of it
      live.batch(() => (plan.cells.C1.value = null))
      inside = [plan.cells.A1.value, plan.cells.A2.value, plan.cells.B1.value, plan.cells.C1.value]
    })
    expect(inside).toEqual([2, 20, null, 'c'])
    expect([plan.cells.A1.text, plan.cells.A2.value, plan.cells.B1.value, plan.cells.C1.value]).toEqual(['7.0', 70, 71, null])
    expect(changes.map((c) => [c.method, c.revision])).toEqual([['applyChanges', 1]])
    expect(session.history.undo).toBe(1)
    session.undo()
    expect([plan.cells.A1.value, plan.cells.B1.value, plan.cells.C1.value]).toEqual([2, null, 'c'])
    // nothing assigned: nothing applied
    live.batch(() => {})
    expect(session.revision).toBe(2)
    session.close()
  })

  it('a batch that throws applies nothing; an inner batch that throws loses its own assignments only', () => {
    const { session, live, plan } = open()
    expect(() =>
      live.batch(() => {
        plan.cells.A1.value = 7
        throw new Error('stop')
      }),
    ).toThrow('stop')
    expect([plan.cells.A1.value, session.revision]).toEqual([2, 0])
    live.batch(() => {
      plan.cells.A1.value = 7
      try {
        live.batch(() => {
          plan.cells.C1.value = 8
          throw new Error('inner')
        })
      } catch {}
    })
    expect([plan.cells.A1.value, plan.cells.C1.value, session.revision]).toEqual([7, 'c', 1])
    // typed text and an fn that answers a promise are not part of a batch
    expect(() => live.batch(() => void plan.cells.A1.type('1'))).toThrow(new InputError('type', 'text typed into a cell cannot be part of a batch (applyChanges takes no typed text): type it before or after the batch'))
    expect(() => live.batch((async () => (plan.cells.A1.value = 9)) as never)).toThrow(new InputError('batch', 'fn must not answer a promise: an assignment after an await is not part of the batch'))
    expect([plan.cells.A1.value, session.revision]).toEqual([7, 1])
    session.close()
  })

  it("a cell's listener is called for the changes that touched it, whoever made them: through the objects, the session, undo and redo; not for others", () => {
    const { session, live, plan } = open()
    const [a1, a2, c1, other] = [heard(), heard(), heard(), heard()]
    plan.cells.A1.watch(a1.listener)
    plan.cells.A2.watch(a2.listener)
    plan.cells.C1.watch(c1.listener)
    live.sheet('Other').cells.A1.watch(other.listener)
    plan.cells.A1.value = 3 // 1: A1, A2 and Other!A1 by the recomputation
    session.setCells([{ sheet: 'Plan', cell: 'A2', input: { value: 1 } }]) // 2: A2 alone
    session.undo() // 3
    session.redo() // 4
    session.patchCells([{ sheet: 'Plan', cell: 'C1', set: { numberFormat: '@' } }]) // 5: C1's look
    session.setLayout([{ sheet: 'Plan', columns: 'C', set: { width: 20 } }]) // 6: nothing a cell shows
    expect([a1.got, a2.got, c1.got, other.got]).toEqual([[1], [1, 2, 3, 4], [5], [1]])
    session.close()
  })

  it("a sheet's listener is called for each change that touched the sheet, the book's for each change", () => {
    const { session, live, plan } = open()
    const [p, o, b] = [heard(), heard(), heard()]
    plan.watch(p.listener)
    live.sheet('Other').watch(o.listener)
    live.watch(b.listener)
    session.setCells([{ sheet: 'Other', cell: 'B1', input: { value: 1 } }]) // 1: Other
    session.setLayout([{ sheet: 'Plan', columns: 'C', set: { width: 20 } }]) // 2: Plan's layout
    session.setStyles([{ font: { bold: true } }]) // 3: the style table, every sheet
    plan.cells.A1.value = 3 // 4: both, Other!A1 depends on Plan!A1
    session.editStructure({ op: 'insertRows', sheet: 'Other', row: 5 }) // 5: Other
    expect([p.got, o.got, b.got]).toEqual([[2, 3, 4], [1, 3, 4, 5], [1, 2, 3, 4, 5]])
    session.close()
  })

  it('one listener of several watches is called once a change, and each watch unsubscribes alone; after unsubscribing it is called no more', () => {
    const { session, live, plan } = open()
    const h = heard()
    const stops = [plan.cells.A1.watch(h.listener), plan.cells.A2.watch(h.listener), plan.watch(h.listener), live.watch(h.listener)]
    plan.cells.A1.value = 3
    expect(h.got).toEqual([1])
    for (const stop of stops.slice(1)) stop()
    plan.cells.A2.value = 3
    plan.cells.A1.value = 4
    stops[0]()
    stops[0]()
    plan.cells.A1.value = 5
    expect(h.got).toEqual([1, 3])
    session.close()
  })

  it('a listener that throws stops neither the edit nor the other listeners; its error is thrown again in a microtask', () => {
    const later: (() => void)[] = []
    vi.stubGlobal('queueMicrotask', (f: () => void) => later.push(f))
    const { session, plan } = open()
    const h = heard()
    plan.cells.A1.watch(() => {
      throw new Error('listener')
    })
    plan.cells.A1.watch(h.listener)
    plan.cells.A1.value = 3
    expect([plan.cells.A2.value, h.got, later.length]).toEqual([30, [1], 1])
    expect(later[0]).toThrow('listener')
    session.close()
  })

  it('a cell keeps its address: after a row is inserted above, it is the cell now there, and its listener is called; those above are not', () => {
    const { session, plan } = open()
    const a5 = plan.cells.A5
    const [h5, h1, h9, c9] = [heard(), heard(), heard(), heard()]
    a5.watch(h5.listener)
    plan.cells.A1.watch(h1.listener)
    plan.cells.A9.watch(h9.listener) // empty before and after
    plan.cells.C9.watch(c9.listener)
    session.editStructure({ op: 'insertRows', sheet: 'Plan', row: 3 })
    expect([a5.address, a5.value, plan.cells.A6.value, plan.cells.A5 === a5]).toEqual(['A5', null, 5, true])
    expect([h5.got, h1.got, h9.got, c9.got]).toEqual([[1], [], [1], [1]])
    session.editStructure({ op: 'insertColumns', sheet: 'Plan', column: 'B' })
    expect([h5.got, h1.got, c9.got]).toEqual([[1], [], [1, 2]])
    // a move wakes where the cells were and where they came to
    const [d1, e1, h7] = [heard(), heard(), heard()]
    plan.cells.D1.watch(d1.listener) // C1 moved by the inserted column
    plan.cells.F2.watch(e1.listener)
    plan.cells.H7.watch(h7.listener)
    session.editStructure({ op: 'moveRange', from: { sheet: 'Plan', range: 'D1:D2' }, to: { sheet: 'Plan', cell: 'F1' } })
    expect([d1.got, e1.got, h7.got, plan.cells.F1.value]).toEqual([[3], [3], [], 'c'])
    session.undo()
    expect([d1.got, e1.got, h7.got, plan.cells.D1.value]).toEqual([[3, 4], [3, 4], [], 'c'])
    session.close()
  })

  it('a sheet follows its rename and ends with its deletion: every use throws, its listeners are called no more, undo makes a new object', () => {
    const { session, live, plan } = open()
    const a1 = plan.cells.A1
    const [h, p] = [heard(), heard()]
    a1.watch(h.listener)
    const stop = plan.watch(p.listener)
    session.editStructure({ op: 'renameSheet', sheet: 'Plan', name: 'Budget' })
    expect([plan.name, live.sheet('budget') === plan, a1.value, a1.sheet === plan, p.got, h.got]).toEqual(['Budget', true, 2, true, [1], []])
    expect(() => live.sheet('Plan')).toThrow(new InputError('name', 'unknown sheet "Plan"'))
    a1.value = 4
    expect([live.sheet('Other').cells.A1.formula, live.sheet('Other').cells.A1.value]).toEqual(["'Budget'!A1+1", 5])
    session.editStructure({ op: 'moveSheet', sheet: 'Budget' })
    expect(p.got).toEqual([1, 2, 3])
    session.editStructure({ op: 'deleteSheet', sheet: 'Budget' })
    const deleted = new InputError('sheet', 'the sheet "Budget" was deleted')
    for (const use of [() => plan.name, () => plan.cells, () => plan.cell('A1'), () => plan.watch(() => {}), () => a1.value, () => (a1.value = 1), () => a1.sheet, () => a1.watch(() => {})]) expect(use).toThrow(deleted)
    expect(stop).not.toThrow()
    session.undo()
    const back = live.sheet('Budget')
    expect([back === plan, back.cells.A1.value, back.cells.A1 === a1, p.got, h.got]).toEqual([false, 4, false, [1, 2, 3], [2]])
    session.close()
  })

  it('cells by address in any case, the same object while held; other keys are undefined, a bad address to cell() throws', () => {
    const { session, live, plan } = open()
    const a1 = plan.cells.A1
    expect([plan.cells.a1 === a1, plan.cell('a1') === a1, plan.cells.XFD1048576.address, 'B7' in plan.cells, 'B7:C8' in plan.cells]).toEqual([true, true, 'XFD1048576', true, false])
    for (const key of ['$A$1', 'A1:B2', 'A0', 'A01', 'XFE1', 'A1048577', 'then', 'toJSON', 'constructor']) expect(plan.cells[key], key).toBe(undefined)
    expect((plan.cells as unknown as Record<symbol, unknown>)[Symbol.iterator]).toBe(undefined)
    expect(JSON.stringify({ c: plan.cells })).toBe('{"c":{}}')
    expect(() => {
      ;(plan.cells as Record<string, unknown>).A1 = 1
    }).toThrow(TypeError)
    for (const bad of ['$A$1', 'A1:B2', '', 5]) expect(() => plan.cell(bad as never)).toThrow(InputError)
    expect(() => plan.cell('A1:B2')).toThrow(new InputError('address', 'expected the address of a cell such as "B12", without $, got the string "A1:B2"'))
    expect(() => live.sheet('Nope')).toThrow(new InputError('name', 'unknown sheet "Nope"'))
    expect(() => plan.cells.A1.watch('x' as never)).toThrow(new InputError('watch', 'listener must be a function, got the string "x"'))
    expect(() => (plan.cells.A1.watch as (...a: unknown[]) => unknown)(() => {}, {})).toThrow(new InputError('watch', 'takes 1 argument, got an object as argument 2'))
    expect([liveBook(session) === live, live.session === session, live.sheet('PLAN') === plan]).toEqual([true, true, true])
    session.close()
  })
})

describe('liveBook takes a synchronous session only', () => {
  it('an asynchronous session or anything else throws InputError at session', async () => {
    const calc = createInlineCalc()
    const s = await calc.createSession(book(), { clock })
    expect(() => liveBook(s as never)).toThrow(new InputError('session', 'expected a synchronous session (createSession), got an asynchronous one (createInlineCalc, createCalcPool)'))
    for (const bad of [null, {}, 'session']) expect(() => liveBook(bad as never)).toThrow(/^session: expected a synchronous session \(createSession\), got /)
    expect(s.revision).toBe(0)
    await s.close()
    await calc.close()
  })
})
