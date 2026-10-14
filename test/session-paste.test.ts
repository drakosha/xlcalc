/**
 * Text pasted from the clipboard (`pasteText`, `previewPasteText`) and a text entered into a selection with
 * Control+Return (`fillCellText`, `previewFillCellText`), in ru-RU. A paste fills the clipboard's rectangle from its cell,
 * each field read as Excel reads it pasted (a number, an apostrophe kept, a formula Excel does not take as text, the
 * euro, no percent editor in a percent cell), a short row and a line break after the last row clearing their cells, no
 * column widened, no rule checked. An entry into a selection reads its text once in the active cell (the top-left one,
 * or another), moves a formula as a copy from it, gives each cell its format, widens the columns and checks the active
 * cell's rule alone. Both in one revision and one step of the undo, refused with the cell where Excel refuses them.
 * In a synchronous session, in the same thread and in a thread, with closures and eval; `parseClipboard` alone.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, InputError, parseClipboard, type AsyncSession, type BookInput, type Execution } from '../src/index.js'
import * as text from '../src/text.js'
import './locales.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'
import { typingParts } from './text-books.js'

afterAll(closeCalcs)

const clock = { now: '2026-10-10T12:00:00Z', seed: 1 }

/** The format Excel gives `€5` pasted. */
const EURO = '[$€-2]\\ #,##0;[Red]\\-[$€-2]\\ #,##0'

const book = (): BookInput => ({
  locale: 'ru-RU',
  styles: [{ font: { bold: true } }],
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 2 }, A2: { value: 46096, numberFormat: 14 }, A3: { value: 3 },
        B2: { value: 'old', style: 0 }, C3: { value: 'old', meta: { id: 7 } }, B4: { value: 'old' },
        C5: { formula: 'A1:A2*10', array: 'C5:C6' },
        E1: { value: 'Qty' }, E2: { value: 1 }, F1: { formula: 'T[Qty]*2' },
      },
      ...typingParts(),
    },
  ],
})

/** Three rows from B1 and a line break after the last: a number with groups, an apostrophe, a percent cell, a formula ru-RU does not take, the euro, a long number, short rows. */
const CLIPBOARD = "1 234,5\t'007\t5\r\n=СУММ(A1;1.5)\t€5\t0.5\r\n1234567890123\r\n"

const cellsOf = async (s: AsyncSession) => (await s.getBook()).sheets[0].cells

/** The book's hash is a full computation's of the book. */
async function computes(s: AsyncSession, execution: Execution) {
  expect(await s.getHash()).toBe(compute(await s.getBook(), { clock, execution }).hash)
}

describe.each(['closures', 'eval'] as const)('a paste and an entry into a selection, execution %s', (execution: Execution) => {
  describe.each(KINDS)('%s', (kind) => {
    const open = async () => (await opener(kind))(book(), { clock, execution })

    it('a paste fills the rectangle, each field read as pasted, clearing what a short row and the last line break leave; no width, no rule; one step', async () => {
      const s = await open()
      const before = structuredClone(await s.getBook())
      const preview = await s.previewPasteText('S', 'B1', CLIPBOARD)
      expect([s.revision, await s.getBook()]).toEqual([0, before])
      const r = await s.pasteText('S', 'B1', CLIPBOARD)
      if (!r.ok || !preview.ok) throw new Error('refused')
      expect([r.range, preview.range, r.revision, s.history]).toEqual(['B1:D4', 'B1:D4', 1, { undo: 1, redo: 0 }])
      const cells = await cellsOf(s)
      expect([cells.B1, cells.C1, cells.D1, cells.B2, cells.C2, cells.D2, cells.B3, cells.C3, cells.D3, cells.B4]).toEqual([
        { value: 1234.5, numberFormat: 4 },
        { value: "'007" },
        // the column's percent format, no percent editor: 500%
        { value: 5 },
        { value: '=СУММ(A1;1.5)', style: 0 },
        { value: 5, numberFormat: EURO },
        { value: '0.5' },
        { value: 1234567890123 },
        // cleared as by empty text: the content goes, the rest stays
        { meta: { id: 7 } },
        undefined,
        undefined,
      ])
      expect(Object.fromEntries(r.changes.map((c) => [c.cell, c.text]))).toEqual({ B1: '1 234,50', C1: "'007", D1: '500%', B2: '=СУММ(A1;1.5)', C2: '€ 5', D2: '0.5', B3: '1,23E+12', C3: '', B4: '' })
      // the preview: every cell of the rectangle with its input and value
      expect(preview.cells.map((c) => c.cell)).toEqual(['B1', 'C1', 'D1', 'B2', 'C2', 'D2', 'B3', 'C3', 'D3', 'B4', 'C4', 'D4'])
      for (const c of preview.cells) expect(c.input, c.cell).toEqual(cells[c.cell] ?? null)
      expect(preview.cells.find((c) => c.cell === 'C2')?.value).toBe(5)
      // no column widened, no rule checked
      expect([r.layout, (await s.getLayout()).sheets[0].columns]).toEqual([undefined, { D: { numberFormat: '0%' } }])
      const ruled = await s.pasteText('S', 'G1', '-2')
      expect(ruled).toMatchObject({ ok: true, range: 'G1' })
      expect(ruled).not.toHaveProperty('validation')
      await computes(s, execution)
      await s.undo()
      const after = structuredClone(await s.getBook())
      const back = await s.undo()
      expect(back).toMatchObject({ call: 'pasteText', revision: 4 })
      expect(await s.getBook()).toEqual(before)
      await s.redo()
      expect(await s.getBook()).toEqual(after)
      await computes(s, execution)
      await s.close()
    })

    it('a paste is refused with the cell over part of an array formula and a formula into a header cell, and past the sheet; a whole array replaced, a header renamed', async () => {
      const s = await open()
      expect(await s.pasteText('S', 'C4', '1\n2')).toEqual({ ok: false, code: 'array', message: 'Нельзя изменить часть массива S!C5 (C5:C6)', sheet: 'S', cell: 'C5' })
      expect(await s.previewPasteText('S', 'D1', '1\t=1+1')).toEqual({ ok: false, code: 'header', message: 'Ячейка S!E1 в строке заголовков таблицы «T»: в ней имя столбца, а не формула', sheet: 'S', cell: 'E1' })
      expect(await s.pasteText('S', 'XFD1048575', '1\n2\n3')).toEqual({ ok: false, code: 'off_sheet', message: 'Копия не помещается на лист «S», начиная с XFD1048575', sheet: 'S', cell: 'XFD1048575' })
      expect([s.revision, s.history.undo, await s.getBook()]).toEqual([0, 0, book()])
      // the whole area: the array goes
      expect(await s.pasteText('S', 'C5', '1\n2')).toMatchObject({ ok: true, range: 'C5:C6' })
      expect([(await cellsOf(s)).C5, (await cellsOf(s)).C6]).toEqual([{ value: 1 }, { value: 2 }])
      // text into a header cell renames its column and the references to it, in the same step
      expect(await s.pasteText('S', 'E1', 'Кол-во\n4')).toMatchObject({ ok: true, range: 'E1:E2' })
      expect([(await cellsOf(s)).F1, (await s.getResult()).sheets[0].cells.F1.value]).toEqual([{ formula: 'T[Кол-во]*2' }, 8])
      await computes(s, execution)
      await s.undo()
      expect((await cellsOf(s)).F1).toEqual({ formula: 'T[Qty]*2' })
      await s.close()
    })

    it('Control+Return: the text read once in the active cell, a formula moved as copied from it, its format in each cell; columns widened; one step', async () => {
      const s = await open()
      const before = structuredClone(await s.getBook())
      const r = await s.fillCellText('S', 'H1:H3', '=A1*2')
      expect(r).toMatchObject({ ok: true, revision: 1 })
      // the active cell elsewhere: references move from it, absolute ones stay
      expect(await s.fillCellText('S', 'I1:J2', '=$A$1+B2', { active: 'J2' })).toMatchObject({ ok: true })
      const cells = await cellsOf(s)
      expect([cells.H1, cells.H2, cells.H3, cells.I1, cells.J1, cells.I2, cells.J2]).toEqual([
        { formula: 'A1*2' }, { formula: 'A2*2' }, { formula: 'A3*2' },
        { formula: '$A$1+A1' }, { formula: '$A$1+B1' }, { formula: '$A$1+A2' }, { formula: '$A$1+B2' },
      ])
      // a percent cell active: 5 is 5% in every cell, which takes its format; a cell in General active: 5 is 5, 500% in the percent cell
      const preview = await s.previewFillCellText('S', 'C1:D1', '5', { active: 'D1' })
      expect(preview.ok && preview.cells.map((c) => [c.cell, c.input, c.value])).toEqual([['C1', { value: 0.05, numberFormat: '0%' }, 0.05], ['D1', { value: 0.05 }, 0.05]])
      await s.fillCellText('S', 'C1:D1', '5')
      expect([(await cellsOf(s)).C1, (await cellsOf(s)).D1]).toEqual([{ value: 5 }, { value: 5 }])
      expect((await s.getCells('S', 'C1:D1')).D1.text).toBe('500%')
      // a number too wide for its columns widens each, as typing does
      const wide = await s.fillCellText('S', 'K1:L2', '1234567890')
      expect(wide).toMatchObject({ ok: true, widths: [{ sheet: 'S', column: 'K', width: 11.1640625 }, { sheet: 'S', column: 'L', width: 11.1640625 }], layout: ['S'] })
      await computes(s, execution)
      for (let i = 0; i < 3; i++) await s.undo()
      expect((await s.undo())?.call).toBe('fillCellText')
      expect(await s.getBook()).toEqual(before)
      await s.close()
    })

    it("Control+Return checks the active cell's rule alone; refused with the cell for a formula Excel does not take, part of an array, a formula into a header cell", async () => {
      const s = await open()
      // the active cell has no rule: G5 and G6 take -2 with no alert
      const quiet = await s.fillCellText('S', 'F5:G6', '-2')
      expect(quiet).toMatchObject({ ok: true })
      expect(quiet).not.toHaveProperty('validation')
      const rule = book().sheets[0].validations![0]
      const loud = await s.fillCellText('S', 'F5:G6', '-3', { active: 'G6' })
      expect(loud).toMatchObject({ ok: true, validation: { valid: false, rule, severity: 'warning', message: 'Positive only' } })
      expect(await s.previewFillCellText('S', 'G1:G2', '-1')).toMatchObject({ ok: true, validation: { severity: 'warning' } })
      expect((await cellsOf(s)).F5).toEqual({ value: -3 })
      const revision = s.revision
      expect(await s.fillCellText('S', 'H1:H2', '=СУММ(A1;', { active: 'H2' })).toEqual({ ok: false, code: 'syntax', message: 'Формула не закончена', at: 9, sheet: 'S', cell: 'H2' })
      expect(await s.fillCellText('S', 'C4:C5', '1')).toEqual({ ok: false, code: 'array', message: 'Нельзя изменить часть массива S!C5 (C5:C6)', sheet: 'S', cell: 'C5' })
      expect(await s.previewFillCellText('S', 'D1:E1', '=1', { active: 'D1' })).toEqual({ ok: false, code: 'header', message: 'Ячейка S!E1 в строке заголовков таблицы «T»: в ней имя столбца, а не формула', sheet: 'S', cell: 'E1' })
      expect(s.revision).toBe(revision)
      await s.close()
    })

    it('arguments: the active cell within the range, a range of a million cells at most, the options each call takes', async () => {
      const s = await open()
      await expect(s.fillCellText('S', 'B1:B3', '1', { active: 'C1' })).rejects.toThrow(new InputError('opts.active', 'expected a cell of the range B1:B3, got "C1"'))
      await expect(s.fillCellText('S', 'A:B', '1')).rejects.toThrow(new InputError('range', 'a range of 2097152 cells: an entry goes into 1048576 cells at most'))
      await expect(s.previewFillCellText('S', 'B1:B0', '1')).rejects.toThrow(InputError)
      await expect(s.pasteText('S', 'B1', 5 as never)).rejects.toThrow('text: expected a string, got number 5')
      await expect(s.pasteText('S', 'B1', '1', { active: 'B1' } as never)).rejects.toThrow('(active is an option of fillCellText and previewFillCellText)')
      expect(s.revision).toBe(0)
      await s.close()
    })
  })
})

describe('parseClipboard', () => {
  it('rows at line breaks, fields at tabs, a rectangle as wide as the longest row; a line break after the last row is a row', () => {
    expect(parseClipboard('1\t2\r\n3\r4\t5\t6\n')).toEqual([['1', '2', ''], ['3', '', ''], ['4', '5', '6'], ['', '', '']])
    expect(parseClipboard('')).toEqual([['']])
  })
  it('a field begun by a quote to its closing quote, tabs and line breaks inside, a doubled quote one, what follows joined; a quote elsewhere a character', () => {
    expect(parseClipboard('"a\tb"')).toEqual([['a\tb']])
    expect(parseClipboard('"x\ny"\tz\n1\t2')).toEqual([['x\ny', 'z'], ['1', '2']])
    expect(parseClipboard('"a""b"c\t"d')).toEqual([['a"bc', 'd']])
    expect(parseClipboard(' "a\nb"')).toEqual([[' "a'], ['b"']])
    expect(text.parseClipboard).toBe(parseClipboard)
    expect(() => parseClipboard(null as never)).toThrow(new InputError('text', 'expected a string, got null'))
  })
})
