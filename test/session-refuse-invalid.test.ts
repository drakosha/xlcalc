/**
 * Typed text refused before it is written where a cell's rule of data validation does not take its value
 * (`refuseInvalid` of `setCellText`, `setCellTexts`, `fillCellText` and their previews): a rule of every style refuses
 * it, with the cell, the rule, its style and message, and the book, the revision and the history stay; a value the rules
 * take, a rule that shows no alert and a cell without a rule are written as without the option, and without it a value
 * is written with its alert as before. Several texts are checked with all of them written and refused whole, with the
 * cell of the strictest rule they break; an entry into a selection checks its active cell alone. In a synchronous session, in the same thread and
 * in a thread, with closures and eval; a refused call replayed after its thread is lost.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { InputError, type AsyncSession, type BookInput, type CellTextOptions, type DataValidation, type Execution } from '../src/index.js'
import { inside } from './fixtures/pool.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const clock = { now: '2026-10-10T12:00:00Z', seed: 1 }

const RULES: DataValidation[] = [
  { range: 'A1:A9', type: 'whole', operator: 'greaterThan', formula1: '0' },
  { range: 'B1:B9', type: 'whole', operator: 'greaterThan', formula1: '0', errorStyle: 'warning', error: 'Positive only' },
  { range: 'C1:C9', type: 'whole', operator: 'greaterThan', formula1: '0', errorStyle: 'information', error: 'Usually positive' },
  { range: 'D1:D9', type: 'whole', operator: 'greaterThan', formula1: '0', showError: false },
  { range: 'G1', type: 'custom', formula1: 'G1<H1' },
]

const book = (): BookInput => ({ sheets: [{ name: 'S', cells: { E1: { formula: 'SUM(A1:C9)' }, H1: { value: 1 } }, validations: RULES }] })

const refuse: CellTextOptions = { refuseInvalid: true }

/** The refusal of a value the rule over `cell` does not take. */
const refusal = (cell: string, rule: DataValidation, severity: string, message = 'The value does not meet the data validation rule of this cell') => ({ ok: false, code: 'validation', message, sheet: 'S', cell, rule, severity })
const STOP = refusal('A1', RULES[0], 'stop')
const WARNING = refusal('B1', RULES[1], 'warning', 'Positive only')
const INFORMATION = refusal('C1', RULES[2], 'information', 'Usually positive')

/** The session changed nothing: its book, revision and history are as it opened. */
async function unchanged(s: AsyncSession) {
  expect(await s.getBook()).toEqual(book())
  expect([s.revision, s.history]).toEqual([0, { undo: 0, redo: 0 }])
}

describe.each(['closures', 'eval'] as const)('typed text refused by a rule of data validation, execution %s', (execution: Execution) => {
  describe.each(KINDS)('%s', (kind) => {
    const open = async () => (await opener(kind))(book(), { clock, execution })

    it('a rule of each style refuses the value with its cell, the rule, its style and message; nothing changes; the host calls again to keep it', async () => {
      const s = await open()
      for (const r of [STOP, WARNING, INFORMATION]) {
        expect(await s.setCellText('S', r.cell, '-2', refuse)).toEqual(r)
        expect(await s.previewCellText('S', r.cell, '-2', refuse)).toEqual(r)
      }
      // a whole number is what each rule takes
      expect(await s.setCellText('S', 'A1', '2.5', refuse)).toEqual(STOP)
      await unchanged(s)
      // the person keeps a warning's value: the call again without the option writes it, with its alert
      const kept = await s.setCellText('S', 'B1', '-2')
      expect(kept).toMatchObject({ ok: true, revision: 1, validation: { valid: false, rule: RULES[1], severity: 'warning', message: 'Positive only' } })
      expect((await s.getBook()).sheets[0].cells.B1).toEqual({ value: -2 })
      expect(s.history).toEqual({ undo: 1, redo: 0 })
      await s.close()
    })

    it('a value the rules take, a rule that shows no alert and a cell without a rule are written as without the option', async () => {
      const [a, b] = [await open(), await open()]
      const calls: ((s: AsyncSession, o?: CellTextOptions) => Promise<unknown>)[] = [
        (s, o) => s.setCellText('S', 'A1', '5', o),
        (s, o) => s.previewCellText('S', 'D1', '-2', o),
        (s, o) => s.setCellText('S', 'D1', '-2', o),
        (s, o) => s.setCellText('S', 'F1', '-2', o),
        (s, o) => s.setCellTexts([{ sheet: 'S', cell: 'B2', text: '3' }, { sheet: 'S', cell: 'C2', text: '4' }], o),
        (s, o) => s.fillCellText('S', 'A3:C4', '7', o),
      ]
      for (const call of calls) expect(await call(a, refuse)).toEqual(await call(b))
      expect(await a.getBook()).toEqual(await b.getBook())
      expect([a.revision, a.history]).toEqual([5, { undo: 5, redo: 0 }])
      await a.close()
      await b.close()
    })

    it("without the option a value no rule takes is written with its alert, and undo leaves a step to redo, as before", async () => {
      const s = await open()
      expect(await s.setCellText('S', 'A1', '-2')).toMatchObject({ ok: true, revision: 1, validation: { valid: false, rule: RULES[0], severity: 'stop', message: STOP.message } })
      expect((await s.getBook()).sheets[0].cells.A1).toEqual({ value: -2 })
      await s.undo()
      expect(s.history).toEqual({ undo: 0, redo: 1 })
      expect((await s.redo())!.changes).toContainEqual({ sheet: 'S', cell: 'A1', value: -2, text: '-2' })
      await s.close()
    })

    it('several texts are checked with all of them written and refused whole, with the cell of the strictest rule they break, the first of those', async () => {
      const s = await open()
      // a stop after a warning: the host that showed the warning and kept the value would write past the stop
      const entries = [{ sheet: 'S', cell: 'F1', text: '7' }, { sheet: 'S', cell: 'C2', text: '-3' }, { sheet: 'S', cell: 'B2', text: '-2' }, { sheet: 'S', cell: 'A2', text: '-1' }, { sheet: 'S', cell: 'A3', text: '-1' }]
      const stop = { ...STOP, cell: 'A2' }
      expect(await s.setCellTexts(entries, refuse)).toEqual(stop)
      expect(await s.previewCellTexts(entries, refuse)).toEqual(stop)
      const softer = [{ sheet: 'S', cell: 'C2', text: '-3' }, { sheet: 'S', cell: 'B2', text: '-2' }, { sheet: 'S', cell: 'B3', text: '-2' }]
      expect(await s.setCellTexts(softer, refuse)).toEqual({ ...WARNING, cell: 'B2' })
      // G1 takes a number below H1: 5 alone it does not take, and with H1 typed to 10 in the same call it does
      expect(await s.setCellTexts([{ sheet: 'S', cell: 'G1', text: '5' }], refuse)).toEqual(refusal('G1', RULES[4], 'stop'))
      await unchanged(s)
      const r = await s.setCellTexts([{ sheet: 'S', cell: 'G1', text: '5' }, { sheet: 'S', cell: 'H1', text: '10' }], refuse)
      expect(r).toMatchObject({ ok: true, revision: 1 })
      expect(r).not.toHaveProperty('validations')
      await s.close()
    })

    it("an entry into a selection checks its active cell's rule alone, as without the option", async () => {
      const s = await open()
      expect(await s.fillCellText('S', 'A3:C4', '-2', refuse)).toEqual({ ...STOP, cell: 'A3' })
      expect(await s.fillCellText('S', 'A3:C4', '-2', { ...refuse, active: 'C4' })).toEqual({ ...INFORMATION, cell: 'C4' })
      expect(await s.previewFillCellText('S', 'A3:C4', '-2', { ...refuse, active: 'B3' })).toEqual({ ...WARNING, cell: 'B3' })
      await unchanged(s)
      // the active cell has no rule: the cells of the others are written, as Excel writes them
      const r = await s.fillCellText('S', 'A5:F5', '-2', { ...refuse, active: 'F5' })
      expect(r).toMatchObject({ ok: true, revision: 1 })
      expect(r).not.toHaveProperty('validation')
      expect((await s.getBook()).sheets[0].cells.A5).toEqual({ value: -2 })
      await s.close()
    })

    it('a live session keeps its clock through a refusal', async () => {
      const s = await (await opener(kind))(book(), { live: true, execution })
      const at = s.clock
      expect(await s.setCellText('S', 'A1', '-2', refuse)).toEqual(STOP)
      expect(s.clock).toEqual(at)
      await s.close()
    })

    it('the option is a boolean, and only the text calls take it', async () => {
      const s = await open()
      await expect(s.setCellText('S', 'A1', '-2', { refuseInvalid: 'yes' } as never)).rejects.toThrow(new InputError('opts.refuseInvalid', 'expected a boolean, got the string "yes"'))
      await expect(s.pasteText('S', 'A1', '-2', refuse as never)).rejects.toThrow('(refuseInvalid is an option of setCellText, setCellTexts, fillCellText and their previews)')
      await unchanged(s)
      await s.close()
    })
  })
})

describe('a thread lost', () => {
  it('a refused call in the journal is replayed as refused: the book, the clock and the history come back as they were', async () => {
    const { calc, crash } = await inside({ threads: 1 })
    try {
      const s = await calc.createSession(book(), { live: true })
      expect(await s.setCellText('S', 'A1', '-2', refuse)).toEqual(STOP)
      await s.setCellText('S', 'A2', '3', refuse)
      expect(await s.setCellTexts([{ sheet: 'S', cell: 'B2', text: '-2' }], refuse)).toEqual({ ...WARNING, cell: 'B2' })
      const [b, at, history] = [await s.getBook(), s.clock, s.history]
      crash(0)
      expect(await s.getBook()).toEqual(b)
      expect([s.clock, s.history, s.revision]).toEqual([at, history, 1])
      expect(await s.fillCellText('S', 'C3:C4', '-2', refuse)).toEqual({ ...INFORMATION, cell: 'C3' })
    } finally {
      await calc.close()
    }
  }, 60_000)
})
