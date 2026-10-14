/**
 * A session's subscribers (`Session.subscribe`, `src/session/subscribers.ts`), in the three kinds of session and both
 * modes of execution: a listener gets each call that changed the book, whoever made it, with what it answered of
 * `EditResult` and its name, before the call returns; not a preview, a read, a refusal or a failed call. Changes reach
 * every listener in the order of their revisions, a change a listener makes too; a listener that throws stops neither
 * the call nor the others; unsubscribing and closing stop the calls.
 */
import { afterAll, afterEach, describe, expect, it, vi } from 'vitest'
import { InputError, type AsyncSession, type BookInput, type EditResult, type SessionChange } from '../src/index.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

const book = (): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { formula: 'A1*10' }, A3: { value: 'x' } } }] })

/** What a change carries of a call's answer: `EditResult`'s fields. */
const edited = (r: EditResult) => {
  const { revision, changes, layout, styles, moved, repaint } = r
  return { revision, changes, ...(layout ? { layout } : {}), ...(styles ? { styles } : {}), ...(moved ? { moved } : {}), ...(repaint ? { repaint } : {}) }
}

/** Every change the session tells, as it is told. */
function heard(s: AsyncSession) {
  const got: SessionChange[] = []
  const stop = s.subscribe((c) => got.push(c))
  return { got, stop }
}

afterAll(closeCalcs)
afterEach(() => vi.unstubAllGlobals())

describe.each(['closures', 'eval'] as const)('subscribers, %s', (execution) => {
  describe.each(KINDS)('%s', (kind) => {
    const open = openerFor(kind)
    const session = () => open(book(), { clock, execution })

    it("each changing call is told with its answer's EditResult and its name, before the call returns", async () => {
      const s = await session()
      const { got } = heard(s)
      // the listener reads the session after the change
      const seen: number[] = []
      s.subscribe(() => seen.push(s.revision))
      const r = await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }])
      expect(got).toEqual([{ method: 'setCells', ...edited(r) }])
      expect(got[0].changes.map((c) => [c.cell, c.value])).toEqual([['A1', 3], ['A2', 30]])
      expect(seen).toEqual([1])
      const t = await s.setCellText('S', 'A3', 'y')
      const moved = await s.editStructure({ op: 'insertRows', sheet: 'S', row: 1 })
      const undone = await s.undo()
      const redone = await s.redo()
      const looks = await s.setLayout([{ sheet: 'S', columns: 'B', set: { width: 20 } }])
      for (const x of [t, moved]) if (!x.ok) throw new Error('refused')
      expect(got.map((c) => c.method)).toEqual(['setCells', 'setCellText', 'editStructure', 'undo', 'redo', 'setLayout'])
      expect(got.slice(1)).toEqual([t, moved, undone!, redone!, looks].map((x, i) => ({ method: got[i + 1].method, ...edited(x as EditResult) })))
      expect(got[2].moved).toEqual([{ kind: 'lines', sheet: 'S', axis: 'rows', at: 1, count: 1 }])
      expect(got.map((c) => c.revision)).toEqual([1, 2, 3, 4, 5, 6])
      await s.close()
    })

    it('a preview, a read, a refusal and a call that throws are not told', async () => {
      const s = await session()
      const { got } = heard(s)
      await s.previewCellText('S', 'A1', '5')
      await s.previewChanges([{ op: 'setCells', edits: [{ sheet: 'S', cell: 'A1', input: { value: 5 } }] }])
      await s.getCells('S', 'A1:A3')
      expect((await s.setCellText('S', 'A1', '=SUM(')).ok).toBe(false)
      await expect(s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 5 } }], { baseRevision: 7 })).rejects.toMatchObject({ code: 'stale_revision' })
      expect(await s.undo()).toBeNull()
      expect([got, s.revision]).toEqual([[], 0])
      await s.close()
    })

    it('a change a listener makes reaches every listener after the one it answers', async () => {
      const s = await session()
      let made: Promise<unknown> | undefined
      const first: number[] = []
      s.subscribe((c) => {
        first.push(c.revision)
        if (c.revision === 1) made = s.setCells([{ sheet: 'S', cell: 'A3', input: { value: 'z' } }])
      })
      const { got } = heard(s)
      await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 4 } }])
      await made
      expect(first).toEqual([1, 2])
      expect(got.map((c) => [c.revision, c.changes.map((x) => x.cell)])).toEqual([[1, ['A1', 'A2']], [2, ['A3']]])
      await s.close()
    })

    it('a listener that throws stops neither the call nor the others; its error comes in a microtask', async () => {
      const thrown: unknown[] = []
      vi.stubGlobal('queueMicrotask', (f: () => void) => {
        try {
          f()
        } catch (e) {
          thrown.push(e)
        }
      })
      const s = await session()
      s.subscribe(() => {
        throw new Error('listener')
      })
      const { got } = heard(s)
      const r = await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 6 } }])
      expect([r.revision, got.length, thrown]).toEqual([1, 1, [new Error('listener')]])
      await s.close()
    })

    it('unsubscribing stops the calls; a listener subscribed twice is called once; a listener must be a function', async () => {
      const s = await session()
      const { got, stop } = heard(s)
      const twice: number[] = []
      const listener = (c: SessionChange) => twice.push(c.revision)
      s.subscribe(listener)
      s.subscribe(listener)
      await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 7 } }])
      stop()
      await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 8 } }])
      expect([got.map((c) => c.revision), twice]).toEqual([[1], [1, 2]])
      expect(() => s.subscribe(5 as never)).toThrow(InputError)
      expect(() => s.subscribe(5 as never)).toThrow('subscribe: listener must be a function, got ')
      expect(() => (s.subscribe as (...a: unknown[]) => unknown)(listener, { once: true })).toThrow(InputError)
      await s.close()
      expect(twice).toEqual([1, 2])
    })
  })
})
