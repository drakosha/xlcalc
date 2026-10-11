/**
 * `viewportOf`, a window over a saved book and its result without a session. The book of `test/fixtures/saved` (two
 * sheets in ru-RU: an error and the cells it goes on to, on the other sheet too, a formula that does not read, a row's
 * and a column's format, a row's style, a merged area, a table, a note, a link, a conditional format, a small spill and
 * a large one whose error the result leaves out) was computed once by xlcalc 0.11.0 and 0.16.3, whose results have the
 * old names (`issues`, a cell's `format`), and by this version; each saved window is compared with a live session's over
 * the same book, with every option, whole and by some rows and columns: all a result holds is the same, and only the
 * conditional formats' appearance, which a result does not hold, is missing. The windows of a session's snapshots of
 * other books, a result without the display, a cycle's origin, the objects of the answer and the input checks. The
 * fixtures are made outside the suite by a script that installs the old versions from their tarballs.
 */
import { readFileSync } from 'node:fs'
import { afterAll, describe, expect, it } from 'vitest'
import { compute, createSession, InputError, viewportOf, type BookInput, type Execution, type SavedResult, type SavedViewport, type Viewport, type ViewportOptions } from '../src/index.js'
import { CLOCK } from './async-books.js'
import { randomBook } from './seeded-books.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'
import { spillBook } from './session-spill-edits.js'
import './locales.js'

afterAll(closeCalcs)

const fixture = <T>(name: string): T => JSON.parse(readFileSync(new URL(`./fixtures/saved/${name}.json`, import.meta.url), 'utf8')) as T
const VERSIONS = ['0.11.0', '0.16.3', '0.17.0']

/** A live window as a saved one gives it: without the session's revision and the conditional formats' appearance. */
function asSaved(v: Viewport): SavedViewport {
  const { revision: _r, ...rest } = v
  for (const c of Object.values(rest.cells)) delete c.appearance
  return rest
}

const ALL: ViewportOptions = { formulaLocal: true, diagnostics: true, editText: true }

/** Windows of the fixture book: whole, with every option, a spill's end, some rows and columns, past `maxCells`. */
const WINDOWS: [string, string, ViewportOptions?][] = [
  ['Данные', 'A1:M10'],
  ['Данные', 'A1:M10', ALL],
  ['Данные', 'J1:L9', { diagnostics: true }],
  ['Данные', 'K69990:L70002', ALL],
  ['Данные', 'A1:L8', { ...ALL, rows: [1, 3, 6, 8], columns: [1, 3, 4, 8, 9, 11] }],
  ['Данные', 'A1:M10', { maxCells: 30 }],
  ['Итоги', 'A1:C4', ALL],
  ['итоги', 'A:A', { maxCells: 5 }],
]

describe.each(['closures', 'eval'] as const)('a saved window against a live one, execution %s', (execution: Execution) => {
  describe.each(KINDS)('%s', (kind) => {
    it('each version\'s result: all it holds as the live window has it, without the appearance', async () => {
      const s = await (await opener(kind))(fixture('book'), { clock: CLOCK, execution })
      for (const version of VERSIONS) {
        const saved = { book: fixture<BookInput>('book'), result: fixture<SavedResult>(`result-${version}`) }
        for (const [sheet, range, opts] of WINDOWS) expect(viewportOf(saved, sheet, range, opts), `${version}: ${sheet}!${range}`).toEqual(asSaved(await s.getViewport(sheet, range, opts)))
      }
      await s.close()
    })

    it("a snapshot's window is the live one, of other books too, without the display as well", async () => {
      const open = await opener(kind)
      const books: [string, BookInput][] = [['the spill book', spillBook()], ['a seeded book', randomBook(5)], ['another', randomBook(11)]]
      for (const [name, book] of books)
        for (const display of [true, false]) {
          const s = await open(book, { clock: CLOCK, execution, display })
          const snapshot = await s.getSnapshot()
          for (const sheet of snapshot.book.sheets)
            for (const range of ['A1:J12', 'B2:D4', 'A5:Z40']) expect(viewportOf(snapshot, sheet.name, range, ALL), `${name}: ${sheet.name}!${range}`).toEqual(asSaved(await s.getViewport(sheet.name, range, ALL)))
          await s.close()
        }
    })
  })
})

describe('a window over a saved book and its result', () => {
  const book = () => fixture<BookInput>('book')

  it('reads the old names of a result, and says what the cells and the errors show', () => {
    const old = fixture<SavedResult>('result-0.11.0')
    expect(old.issues).toBeDefined()
    expect(old.diagnostics).toBeUndefined()
    // a date that its column is too narrow for, as Excel shows it
    expect(old.sheets[0].cells.A8).toEqual({ value: 46096, text: '#########', format: 14 })
    const v = viewportOf({ book: book(), result: old }, 'Данные', 'A1:L8', ALL)
    expect(v.cells.A8).toEqual({ value: 46096, text: '#########', numberFormat: 'dd.mm.yyyy', numberFormatSource: 'cell', editText: '15.03.2026' })
    // a rule's appearance is not in a result: the cell as the result has it
    expect(v.cells.C3).toEqual({ value: 20, text: '20', editText: '20' })
    expect(v.cells.D2).toMatchObject({ formulaLocal: '[@Кол]*[@Цена]', editText: '=[@Кол]*[@Цена]' })
    expect(v.cells.F3.issue).toEqual({ error: 'DIV0', reason: "division by zero, from 'Данные'!F1", origin: { sheet: 'Данные', cell: 'F1' } })
    // through a cell of a large spill that the result leaves out, to its formula
    expect(v.cells.L1.issue).toEqual({ error: 'DIV0', reason: "division by zero, from 'Данные'!K5", origin: { sheet: 'Данные', cell: 'K1' } })
    expect(v.cells.K5).toMatchObject({ value: { error: 'DIV0' }, text: '#ДЕЛ/0!', spillOf: 'K1', issue: { reason: "division by zero, in the array spilled from 'Данные'!K1", origin: { sheet: 'Данные', cell: 'K1' } } })
    expect(v.cells.F4.issue).toEqual({ error: 'ERROR', reason: 'unexpected end of formula', check: 'syntax', at: 4, origin: { sheet: 'Данные', cell: 'F4' } })
    expect(viewportOf({ book: book(), result: old }, 'Итоги', 'A1', ALL).cells.A1.issue?.origin).toEqual({ sheet: 'Данные', cell: 'F1' })
    expect(v.usedRange).toEqual({ rows: 70_000, columns: 12 })
    expect(v.merges).toEqual(['H6:I7'])
    expect(v.tables).toEqual([{ name: 'Заказы', range: 'A1:D3', headers: true, totals: false }])
  })

  it('a result without the display: values alone, as a session without it gives them', () => {
    const b = book()
    const v = viewportOf({ book: b, result: compute(b, { clock: CLOCK, display: false }) }, 'Данные', 'A5:C6')
    expect(v.cells).toEqual(asSaved(createSession(book(), { clock: CLOCK, display: false }).getViewport('Данные', 'A5:C6')).cells)
    expect(v.cells.C6).toEqual({ value: null, effectiveStyle: 0 })
  })

  it("a cycle's reasons name no cell: its cells' errors begin where they are", () => {
    const b: BookInput = { sheets: [{ name: 'S', cells: { A1: { formula: 'B1' }, B1: { formula: 'A1' }, C1: { formula: 'A1+1' } } }] }
    const v = viewportOf({ book: b, result: compute(b, { clock: CLOCK }) }, 'S', 'A1:C1', { diagnostics: true })
    expect(Object.values(v.cells).map((c) => [c.issue?.reason, c.issue?.origin.cell])).toEqual([['circular reference', 'A1'], ['circular reference', 'B1'], ['circular reference', 'C1']])
  })

  it('the objects are read once and kept with them; the answer has its own', () => {
    const saved = { book: book(), result: fixture<SavedResult>('result-0.17.0') }
    const first = viewportOf(saved, 'Данные', 'A1:I8')
    first.usedRange.rows = 1
    first.styles[0].font!.bold = false
    const again = viewportOf(saved, 'Данные', 'A1:I8')
    expect(again).toMatchObject({ usedRange: { rows: 70_000, columns: 12 }, styles: { 0: { font: { bold: true } } } })
    expect(viewportOf({ ...saved }, 'Данные', 'A1:I8')).toEqual(again)
  })

  it('the book, the result, the sheet, the range and the options are checked as input', () => {
    const result = fixture<SavedResult>('result-0.17.0')
    const call = (saved: unknown, sheet = 'Данные', range = 'A1', opts?: unknown) => () => viewportOf(saved as never, sheet, range, opts as never)
    expect(call(null)).toThrow(InputError)
    expect(call({ book: { sheets: 'x' }, result })).toThrow('saved.book.sheets')
    expect(call({ book: book(), result: { ...result, sheets: result.sheets.slice(1) } })).toThrow('saved.result.sheets: expected the 2 sheets of the book, got 1: the result is not of this book')
    expect(call({ book: book(), result: { ...result, sheets: [...result.sheets].reverse() } })).toThrow('saved.result.sheets[0].name')
    expect(call({ book: book(), result: { ...result, clock: undefined } })).toThrow('saved.result.clock')
    expect(call({ book: book(), result }, 'Нет')).toThrow('sheet: unknown sheet "Нет"')
    expect(call({ book: book(), result }, 'Данные', 'A0')).toThrow('range: expected a range')
    expect(call({ book: book(), result }, 'Данные', 'A1', { maxcells: 5 })).toThrow(InputError)
    expect(call({ book: book(), result }, 'Данные', 'A1:B9', { rows: [2, 2] })).toThrow('opts.rows[1]: expected numbers in ascending order, each once: 2 comes after 2')
  })
})
