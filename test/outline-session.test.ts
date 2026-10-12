/**
 * Excel's outline commands in a session (`editStructure` with `group`, `ungroup`, `showDetail`, `hideDetail`,
 * `showLevels`): one revision and one step of the undo each, SUBTOTAL 109 recomputed over the rows they hide and show,
 * the sheet's groups after each as `outlineGroups` gives them, refusals as Excel's with a message in the book's locale, in
 * a synchronous session, in the same thread and in a thread, with closures and eval.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, outlineGroups, type BookInput, type Execution, type OutlineGroup } from '../src/index.js'
import './locales.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const clock = { now: '2026-10-12T12:00:00Z', seed: 1 }

/** Rows 3:10 grouped (level 1, summary row 11) and 5:7 inside (level 2, summary row 8), each row holding its number; SUBTOTALs over them. */
const book = (locale?: string): BookInput => ({
  ...(locale ? { locale } : {}),
  sheets: [
    {
      name: 'S',
      cells: { ...Object.fromEntries([3, 4, 5, 6, 7, 8, 9, 10].map((r) => [`A${r}`, { value: r }])), B1: { formula: 'SUBTOTAL(109,A3:A10)' }, B2: { formula: 'SUBTOTAL(9,A3:A10)' } },
      rows: { '3:4': { outlineLevel: 1 }, '5:7': { outlineLevel: 2 }, '8:10': { outlineLevel: 1 } },
    },
  ],
})

/** A group as `[first, last, level, collapsed]`. */
const short = (gs: OutlineGroup[]) => gs.map((g) => [g.first, g.last, g.level, g.collapsed])

describe.each(['closures', 'eval'] as Execution[])('outline commands in a session, execution %s', (execution) => {
  describe.each(KINDS)('%s', (kind) => {
    it("Hide and Show Detail as Excel: the inner group stays collapsed when the outer one opens; one revision and one step of the undo each; SUBTOTAL 109 follows", async () => {
      const s = await (await opener(kind))(book(), { clock, execution })
      const groups = async () => short(outlineGroups((await s.getLayout()).sheets[0]).rows)
      const sum = async () => (await s.getCell('S', 'B1'))?.value
      let r = await s.editStructure({ op: 'hideDetail', sheet: 'S', row: 8 })
      expect(r.ok && [r.revision, r.changes.map((c) => [c.cell, c.value]), r.mapCell({ sheet: 'S', cell: 'A5' })]).toEqual([1, [['B1', 34]], { sheet: 'S', cell: 'A5' }])
      expect(await groups()).toEqual([[3, 10, 1, false], [5, 7, 2, true]])
      r = await s.editStructure({ op: 'hideDetail', sheet: 'S', row: 4 })
      expect(r.ok && r.changes.map((c) => [c.cell, c.value])).toEqual([['B1', 0]])
      expect(await groups()).toEqual([[3, 10, 1, true], [5, 7, 2, true]])
      // the grid's example: rows 3, 4, 8, 9 and 10 show; 5:7 stay hidden with the mark on row 8; row 11 loses its mark
      r = await s.editStructure({ op: 'showDetail', sheet: 'S', row: 11 })
      expect(r.ok && r.changes.map((c) => [c.cell, c.value])).toEqual([['B1', 34]])
      expect((await s.getLayout()).sheets[0].rows).toEqual({ '3:4': { outlineLevel: 1 }, '5:7': { outlineLevel: 2, hidden: true }, 8: { outlineLevel: 1, collapsed: true }, '9:10': { outlineLevel: 1 } })
      expect([s.revision, s.history]).toEqual([3, { undo: 3, redo: 0 }])
      // refused as Excel refuses: the detail shown already, nothing changes
      expect(await s.editStructure({ op: 'showDetail', sheet: 'S', row: 11 })).toEqual({ ok: false, code: 'detail', message: 'the detail of row 11 is shown already' })
      expect(s.revision).toBe(3)
      const u = await s.undo()
      expect(u && [u.call, u.changes.map((c) => [c.cell, c.value])]).toEqual(['editStructure', [['B1', 0]]])
      expect(await groups()).toEqual([[3, 10, 1, true], [5, 7, 2, true]])
      await s.redo()
      expect([await groups(), await sum()]).toEqual([[[3, 10, 1, false], [5, 7, 2, true]], 34])
      expect(await s.getHash()).toBe(compute(await s.getBook(), { clock, execution }).hash)
      await s.close()
    })

    it('Group, Ungroup and the level buttons; their refusals in the book\'s locale', async () => {
      const s = await (await opener(kind))(book('ru-RU'), { clock, execution })
      const groups = async () => short(outlineGroups((await s.getLayout()).sheets[0]).rows)
      expect((await s.editStructure({ op: 'group', sheet: 'S', rows: '11:12' })).ok).toBe(true)
      // rows 3:12 one group of level 1 now
      expect(await groups()).toEqual([[3, 12, 1, false], [5, 7, 2, false]])
      let r = await s.editStructure({ op: 'showLevels', sheet: 'S', rows: 2 })
      expect(r.ok && r.changes.map((c) => [c.cell, c.value])).toEqual([['B1', 34]])
      expect(await groups()).toEqual([[3, 12, 1, false], [5, 7, 2, true]])
      r = await s.editStructure({ op: 'showLevels', sheet: 'S', rows: 1 })
      expect(r.ok && r.changes.map((c) => [c.cell, c.value])).toEqual([['B1', 0]])
      expect(await groups()).toEqual([[3, 12, 1, true], [5, 7, 2, true]])
      // ungrouped, the hidden rows stay hidden
      expect((await s.editStructure({ op: 'ungroup', sheet: 'S', rows: '3:12' })).ok).toBe(true)
      expect([await groups(), await s.getCell('S', 'B1').then((c) => c?.value)]).toEqual([[[5, 7, 1, true]], 0])
      expect(await s.editStructure({ op: 'ungroup', sheet: 'S', columns: 'A:C' })).toEqual({ ok: false, code: 'not_grouped', message: 'Столбцы A:C не входят ни в одну группу' })
      expect(await s.editStructure({ op: 'hideDetail', sheet: 'S', row: 20 })).toEqual({ ok: false, code: 'not_grouped', message: 'Строка 20 не входит в группу с итоговой строкой' })
      for (let i = 0; i < 6; i++) await s.editStructure({ op: 'group', sheet: 'S', rows: '5:7' })
      expect(await s.editStructure({ op: 'group', sheet: 'S', rows: '6' })).toEqual({ ok: false, code: 'outline_deep', message: 'Нельзя сгруппировать строки 6: они уже на самом глубоком уровне структуры, 7' })
      expect(s.history.undo).toBe(10)
      await expect(s.editStructure({ op: 'showLevels', sheet: 'S', rows: 9 })).rejects.toThrow('op.rows: expected an integer from 1 to 8, got 9')
      await expect(s.editStructure({ op: 'hideDetail', sheet: 'S', rows: '5:7' } as never)).rejects.toThrow('unknown field "rows"')
      await s.close()
    })
  })
})
