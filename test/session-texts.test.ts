/**
 * Texts typed into cells at once (`setCellTexts`, `previewCellTexts`): the book, the values and the changes are those of
 * the texts typed one by one with `setCellText` (numbers, dates and a formula in ru-RU, text that reads as a number, a
 * column's format, a cell with a rule of data validation, two cells of a column that widen it to different widths, a cell
 * given twice, a table's header typed over), in one revision and one step of the undo; a column widened once to its
 * widest, or the widths only told; the cells whose values their rules do not take; the first text refused refuses all,
 * with its cell. In a synchronous session, in the same thread and in a thread, with closures and eval.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, InputError, type AsyncSession, type BookInput, type CellChange, type CellTextEntry, type Execution } from '../src/index.js'
import './locales.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'
import { typingParts } from './text-books.js'

afterAll(closeCalcs)

const clock = { now: '2026-10-10T12:00:00Z', seed: 1 }

const book = (): BookInput => ({
  locale: 'ru-RU',
  styles: [{ font: { bold: true } }],
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 1, style: 0, note: { text: 'n' } },
        B1: { formula: 'SUM(A1:A9)' },
        C5: { formula: 'A1:A2*10', array: 'C5:C6' },
        E1: { value: 'Qty' },
        E2: { value: 1 },
        F1: { formula: 'T[Qty]*2' },
      },
      ...typingParts(),
    },
    { name: 'Q', cells: { A1: { formula: "S!A3+1" } } },
  ],
})

/** Numbers, dates and a formula in ru-RU, text that reads as a number, a column's format, a rule, a column widened by two cells, another sheet, a cell given twice. */
const MIXED: CellTextEntry[] = [
  { sheet: 'S', cell: 'A2', text: '1 234,5' },
  { sheet: 'S', cell: 'A3', text: '15.03.2026' },
  { sheet: 'S', cell: 'A4', text: '1234567890' },
  { sheet: 'S', cell: 'C1', text: '=СУММ(A1:A4;0,5)' },
  { sheet: 'S', cell: 'C2', text: "'007" },
  { sheet: 'S', cell: 'D1', text: '5' },
  { sheet: 'S', cell: 'G1', text: '-3' },
  { sheet: 'Q', cell: 'B2', text: '12345678901' },
  { sheet: 'S', cell: 'A1', text: '7' },
  { sheet: 'S', cell: 'A7', text: '15.03.2026' },
  { sheet: 'S', cell: 'A7', text: '45' },
]

/** The changes of several answers as one: each cell as the last answer that lists it has it, with the fields earlier ones gave. */
function merged(answers: { changes: CellChange[] }[]): Record<string, CellChange> {
  const out: Record<string, CellChange> = {}
  for (const a of answers) for (const c of a.changes) out[`${c.sheet}!${c.cell}`] = { ...out[`${c.sheet}!${c.cell}`], ...c }
  return out
}

const byCell = (changes: CellChange[]) => Object.fromEntries(changes.map((c) => [`${c.sheet}!${c.cell}`, c]))

/** The texts typed one by one into a session of its own, every answer taken. */
async function oneByOne(s: AsyncSession, entries: CellTextEntry[]) {
  const answers = []
  for (const e of entries) {
    const r = await s.setCellText(e.sheet, e.cell, e.text)
    if (!r.ok) throw new Error(r.message)
    answers.push(r)
  }
  return answers
}

describe.each(['closures', 'eval'] as const)('texts typed into cells at once, execution %s', (execution: Execution) => {
  describe.each(KINDS)('%s', (kind) => {
    it('the book, the values and the changes of the texts typed one by one, in one revision and one step of the undo', async () => {
      const open = await opener(kind)
      const [one, all] = [await open(book(), { clock, execution }), await open(book(), { clock, execution })]
      const answers = await oneByOne(one, MIXED)
      const r = await all.setCellTexts(MIXED)
      if (!r.ok) throw new Error(r.message)
      expect(await all.getBook()).toEqual(await one.getBook())
      expect(await all.getHash()).toBe(await one.getHash())
      expect((await all.getResult()).sheets).toEqual((await one.getResult()).sheets)
      expect(byCell(r.changes)).toEqual(merged(answers))
      // column A widened once, to the widest of its cells (a date, then a longer number one by one); C by the formula's number, Q!B by its own; the value the rule does not take, with its cell
      expect(r.widths).toEqual([
        { sheet: 'S', column: 'A', width: 11.1640625 },
        { sheet: 'S', column: 'C', width: 11.1640625 },
        { sheet: 'Q', column: 'B', width: 12.1640625 },
      ])
      expect(answers.flatMap((a) => (a.width === undefined ? [] : [`${a.changes[0].cell} ${a.width}`]))).toEqual(['A3 10.1640625', 'A4 11.1640625', 'C1 11.1640625', 'B2 12.1640625'])
      const rule = book().sheets[0].validations![0]
      expect(r.validations).toEqual([{ sheet: 'S', cell: 'G1', valid: false, rule, severity: 'warning', message: 'Positive only' }])
      expect(r.layout).toEqual(['S', 'Q'])
      expect([r.revision, all.revision, all.history]).toEqual([1, 1, { undo: 1, redo: 0 }])
      const cells = (await all.getBook()).sheets[0].cells
      expect([cells.A1, cells.A2, cells.A3, cells.C1, cells.C2, cells.D1, cells.A7]).toEqual([
        { value: 7, style: 0, note: { text: 'n' } },
        { value: 1234.5, numberFormat: 4 },
        { value: 46096, numberFormat: 14 },
        { formula: 'SUM(A1:A4,0.5)' },
        { value: '007' },
        { value: 0.05 },
        { value: 45, numberFormat: 14 },
      ])
      expect(await all.getHash()).toBe(compute(await all.getBook(), { clock, execution }).hash)
      await Promise.all([one.close(), all.close()])
    })

    it('one undo takes the whole call back, widths too; redo does it again', async () => {
      const s = await (await opener(kind))(book(), { clock, execution })
      const before = structuredClone(await s.getBook())
      const r = await s.setCellTexts(MIXED)
      const after = structuredClone(await s.getBook())
      const back = await s.undo()
      expect(back).toMatchObject({ call: 'setCellTexts', revision: 2, layout: ['S', 'Q'] })
      expect(await s.getBook()).toEqual(before)
      expect(await s.getHash()).toBe(compute(before, { clock, execution }).hash)
      expect(s.history).toEqual({ undo: 0, redo: 1 })
      const forth = await s.redo()
      expect(await s.getBook()).toEqual(after)
      expect(byCell(forth!.changes)).toEqual(byCell(r.ok ? r.changes : []))
      await s.close()
    })

    it("a header typed over renames its column with the other texts and their widths, as one by one; undone and done again", async () => {
      const entries: CellTextEntry[] = [
        { sheet: 'S', cell: 'E1', text: 'Кол-во' },
        { sheet: 'S', cell: 'H1', text: '1234567890' },
        { sheet: 'S', cell: 'E2', text: '4' },
      ]
      const open = await opener(kind)
      const [one, all] = [await open(book(), { clock, execution }), await open(book(), { clock, execution })]
      const before = structuredClone(await all.getBook())
      const answers = await oneByOne(one, entries)
      const r = await all.setCellTexts(entries)
      expect(r).toMatchObject({ ok: true, widths: [{ sheet: 'S', column: 'H', width: 11.1640625 }] })
      expect(await all.getBook()).toEqual(await one.getBook())
      expect((await all.getBook()).sheets[0].cells.F1).toEqual({ formula: 'T[Кол-во]*2' })
      expect(byCell(r.ok ? r.changes : [])).toEqual(merged(answers))
      const after = structuredClone(await all.getBook())
      await all.undo()
      expect(await all.getBook()).toEqual(before)
      await all.redo()
      expect(await all.getBook()).toEqual(after)
      expect(await all.getHash()).toBe(compute(after, { clock, execution }).hash)
      await Promise.all([one.close(), all.close()])
    })

    it('widen: false tells the widths and sets none; the preview tells the inputs, values, widths and alerts with nothing written', async () => {
      const s = await (await opener(kind))(book(), { clock, execution })
      const preview = await s.previewCellTexts(MIXED)
      expect(s.revision).toBe(0)
      expect(await s.getBook()).toEqual(book())
      const r = await s.setCellTexts(MIXED, { widen: false })
      if (!r.ok || !preview.ok) throw new Error('refused')
      expect(r.widths).toEqual(preview.widths)
      expect(r.validations).toEqual(preview.validations)
      expect(r.layout).toBeUndefined()
      expect((await s.getLayout()).sheets.map((x) => x.columns)).toEqual([{ D: { numberFormat: '0%' } }, undefined])
      const b = await s.getBook()
      const result = await s.getResult()
      const sheets = new Map(b.sheets.map((x, i) => [x.name, [x.cells, result.sheets[i].cells] as const]))
      expect(preview.cells.map((c) => c.cell)).toEqual(['A2', 'A3', 'A4', 'C1', 'C2', 'D1', 'G1', 'B2', 'A1', 'A7'])
      for (const c of preview.cells) expect({ input: c.input, value: c.value }, `${c.sheet}!${c.cell}`).toEqual({ input: sheets.get(c.sheet)![0][c.cell], value: sheets.get(c.sheet)![1][c.cell].value })
      await s.close()
    })

    it('the first text refused refuses them all, with its cell: nothing written, no revision, no step', async () => {
      const s = await (await opener(kind))(book(), { clock, execution })
      const typed = (cell: string, text: string): CellTextEntry[] => [{ sheet: 'S', cell: 'A2', text: '5' }, { sheet: 'S', cell, text }, { sheet: 'S', cell: 'A3', text: '=СУММ(' }]
      expect(await s.setCellTexts(typed('B3', '=СУММ(A1;1.5)'))).toEqual({ ok: false, code: 'syntax', message: 'Здесь не может стоять «.»', at: 10, sheet: 'S', cell: 'B3' })
      expect(await s.previewCellTexts(typed('B3', '=СУММ(A1;1.5)'))).toEqual({ ok: false, code: 'syntax', message: 'Здесь не может стоять «.»', at: 10, sheet: 'S', cell: 'B3' })
      expect(await s.setCellTexts(typed('C6', '7'))).toEqual({ ok: false, code: 'array', message: 'Нельзя изменить часть массива S!C5 (C5:C6)', sheet: 'S', cell: 'C6' })
      expect(await s.setCellTexts(typed('E1', '=1+1'))).toEqual({ ok: false, code: 'header', message: 'Ячейка S!E1 в строке заголовков таблицы «T»: в ней имя столбца, а не формула', sheet: 'S', cell: 'E1' })
      expect(await s.getBook()).toEqual(book())
      expect([s.revision, s.history.undo]).toEqual([0, 0])
      // nothing to type is no step
      expect(await s.setCellTexts([])).toEqual({ ok: true, revision: 1, changes: [] })
      expect(s.history.undo).toBe(0)
      await s.close()
    })

    it('entries of the wrong shape are an InputError naming the entry; an option it does not take too', async () => {
      const s = await (await opener(kind))(book(), { clock, execution })
      await expect(s.setCellTexts({ sheet: 'S' } as never)).rejects.toThrow(new InputError('entries', 'expected an array of texts such as [{ sheet: "Sheet1", cell: "A1", text: "15.03.2026" }], got an object'))
      await expect(s.setCellTexts([{ sheet: 'S', cell: 'A1', text: '1' }, { sheet: 'S', cell: 'A2' } as never])).rejects.toThrow('entries[1].text: expected a string, got undefined')
      await expect(s.previewCellTexts([{ sheet: 'Nope', cell: 'A1', text: '1' }])).rejects.toThrow('entries[0].sheet: unknown sheet "Nope"')
      await expect(s.setCellTexts([{ sheet: 'S', cell: 'a1', text: '1' }])).rejects.toThrow(InputError)
      await expect(s.setCellTexts([{ sheet: 'S', cell: 'A1', text: '1', input: {} } as never])).rejects.toThrow('entries[0]: unknown field "input" (expected sheet, cell, text)')
      await expect(s.setCellTexts([], { maxCells: 2 } as never)).rejects.toThrow('(maxCells is an option of getViewport)')
      expect(s.revision).toBe(0)
      await s.close()
    })
  })
})
