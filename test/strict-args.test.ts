/**
 * What a call does not take is refused, not ignored (`src/strict-args.ts`): an options object or a callback past a
 * method's or a function's own arguments is an `InputError` naming the call that takes those options, and so is an
 * unknown field of a session call's options; an array, a number or undefined past them passes, as `map` gives a callback
 * its index and array. Every method of the three kinds of session and of both calculators, every function of the root,
 * `xlcalc/a1`, `xlcalc/xlsx` and `xlcalc/worker`. With closures and eval.
 */
import { afterAll, describe, expect, it } from 'vitest'
import * as root from '../src/index.js'
import { compute, createInlineCalc, createSession, InputError, parseCell, toInvariant, toLocal, type AsyncSession, type BookInput, type Execution } from '../src/index.js'
import * as a1 from '../src/a1.js'
import * as text from '../src/text.js'
import * as xlsx from '../src/xlsx.js'
import * as worker from '../src/worker/index.js'
import { ruRU } from '../src/locales/ru-RU.js'
import { CLOCK } from './async-books.js'
import { closeCalcs, keptPool, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const book = (): BookInput => ({ styles: [{ font: { bold: true } }], sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { formula: 'A1*21' } } }] })

/** The arguments each method of a synchronous session takes (`Session`). */
const SESSION: Record<string, number> = {
  applyChanges: 2, previewChanges: 2, readCells: 2, findCells: 3, replaceInCells: 4, getDependents: 2, summarizeRanges: 2,
  setCells: 2, replaceCells: 2, clearCells: 3, patchCells: 2, setCellText: 4, previewCellText: 4, setCellTexts: 2, previewCellTexts: 2, pasteText: 4, previewPasteText: 4, fillCellText: 4, previewFillCellText: 4, setNames: 2, setClock: 2, editStructure: 2, setStyles: 2, setLayout: 2,
  getBook: 0, getLayout: 0, getUsedRange: 1, getEnd: 3, explain: 3, getCells: 2, copyText: 3, getCell: 2, getViewport: 3, getCellIssue: 2, getValidation: 2, validateEntry: 3,
  getResult: 0, getSnapshot: 0, getHash: 0, undo: 1, redo: 1, close: 0,
}

/** An asynchronous session's (`AsyncSession`): the calls that compute take a call's options too. */
const ASYNC: Record<string, number> = { ...SESSION, validateEntry: 4, getResult: 1, getSnapshot: 1, getHash: 1 }

const GETCELLS = new InputError('getCells', 'takes 2 arguments, got an object as argument 3 (formulaLocal is an option of getViewport)')

describe('a session takes no options a method does not', () => {
  it("each method of a synchronous session by its parameters, and of an asynchronous one by its own", async () => {
    const s = createSession(book(), { clock: CLOCK })
    const own = (o: object) => Object.fromEntries(Object.entries(o).filter(([, f]) => typeof f === 'function').map(([k, f]) => [k, (f as () => void).length]))
    expect(own(s)).toEqual(SESSION)
    const inline = createInlineCalc()
    const pool = await keptPool()
    for (const calc of [inline, pool]) {
      const a = await calc.createSession(book(), { clock: CLOCK })
      expect({ ...own(a), kill: undefined, historyData: undefined }).toEqual({ ...ASYNC, kill: undefined, historyData: undefined })
      expect(own(calc)).toEqual({ compute: 2, createSession: 2, close: 0 })
      await a.close()
    }
    await inline.close()
    s.close()
  })
})

describe.each(['closures', 'eval'] as const)('arguments a call does not take, execution %s', (execution: Execution) => {
  describe.each(KINDS)('%s', (kind) => {
    const open = async (): Promise<AsyncSession> => (await opener(kind))(book(), { clock: CLOCK, execution })
    const table = kind === 'a synchronous session' ? SESSION : ASYNC

    it("getCells with getViewport's option: refused, saying whose option it is; the session as it was", async () => {
      const s = await open()
      // the call a JavaScript host can make, which the types refuse
      const getCells = s.getCells as (...a: unknown[]) => Promise<object>
      await expect(getCells('S', 'A1:A2', { formulaLocal: true })).rejects.toThrow(GETCELLS)
      await expect(getCells('S', 'A1:A2', { formulaLocal: true })).rejects.toBeInstanceOf(InputError)
      // undefined, a number or an array past the arguments pass, as map gives them
      expect(await getCells('S', 'A2', undefined)).toEqual({ A2: { value: 42, text: '42', formula: 'A1*21' } })
      expect(await getCells('S', 'A2', 0, [])).toEqual({ A2: { value: 42, text: '42', formula: 'A1*21' } })
      expect([s.revision, (await s.getResult()).sheets[0].cells.A2.value]).toEqual([0, 42])
      await s.close()
    })

    it('every method refuses an options object or a callback past its arguments, by its name', async () => {
      const s = await open()
      for (const [name, max] of Object.entries(table)) {
        if (name === 'close') continue
        const call = (s as unknown as Record<string, (...a: unknown[]) => Promise<unknown>>)[name]
        await expect(call(...Array(max).fill(undefined), { depth: 2 }), name).rejects.toThrow(`${name}: takes ${max} argument${max === 1 ? '' : 's'}, got an object as argument ${max + 1} (depth is an option of explain)`)
        await expect(call(...Array(max).fill(undefined), () => 0), name).rejects.toThrow(InputError)
      }
      await expect((s.close as (...a: unknown[]) => Promise<void>)({ signal: null })).rejects.toThrow('close: takes 0 arguments, got an object as argument 1')
      await s.close()
    })

    it("an unknown option of a session's call is refused with the call that takes it", async () => {
      const s = await open()
      await expect(s.explain('S', 'A2', { formulaLocal: true } as never)).rejects.toThrow('opts: unknown field "formulaLocal" (expected depth, maxNodes, expand, errorDepth) (formulaLocal is an option of getViewport)')
      await expect(s.getViewport('S', 'A1', { depth: 2 } as never)).rejects.toThrow('(depth is an option of explain)')
      // a session's window has the appearance always: the option is the saved window's
      await expect(s.getViewport('S', 'A1', { appearance: true } as never)).rejects.toThrow('(appearance is an option of viewportOf)')
      await expect(s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }], { formulaLocal: true } as never)).rejects.toThrow('(formulaLocal is an option of getViewport)')
      await expect(s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }], { clock: CLOCK } as never)).rejects.toThrow('(clock is an option of compute and createSession)')
      await expect(s.setCellText('S', 'A1', '3', { widen: false, maxCells: 2 } as never)).rejects.toThrow('(maxCells is an option of getViewport)')
      await expect(s.pasteText('S', 'A1', '3', { widen: false } as never)).rejects.toThrow('(widen is an option of setCellText, setCellTexts, fillCellText and their previews)')
      await expect(s.setCellText('S', 'A1', '3', { active: 'A1' } as never)).rejects.toThrow('(active is an option of fillCellText and previewFillCellText)')
      expect((await s.explain('S', 'A2', { depth: 1, errorDepth: 2 })).text).toContain('A1*21')
      expect(s.revision).toBe(0)
      await s.close()
    })

    it("a style by its fields where a cell takes its number says where such a style goes", async () => {
      const s = await open()
      await expect(s.setCells([{ sheet: 'S', cell: 'A1', input: { style: { font: { bold: true } } as never } }])).rejects.toThrow("edits[0].input.style: expected the index of a style of the book, 0 to 0, got an object (a cell names a style of the book's styles by its index; a session's patchCells takes one by its fields)")
      expect((await s.patchCells([{ sheet: 'S', cell: 'A1', set: { style: { font: { bold: true } } } }])).changes).toEqual([{ sheet: 'S', cell: 'A1', value: 2, text: '2', style: 0 }])
      await s.close()
    })
  })
})

describe('the functions of the package take no options they do not', () => {
  /** Every function the package gives out, by entry point, and the arguments it takes. */
  const FUNCTIONS: [string, Record<string, unknown>, Record<string, number>][] = [
    ['xlcalc', root, {
      compute: 2, createSession: 2, createInlineCalc: 1, viewportOf: 4, editStructure: 2, copyFormula: 4, validateFormula: 2, validateLocalFormula: 3, validateName: 2, validateNames: 2,
      toLocal: 3, toInvariant: 3, registerLocale: 1, parseEntry: 2, entryText: 2, needsApostrophe: 2, entryWidth: 2, fitColumns: 4, fitRows: 4, displayValue: 2, builtinFormatCode: 2, numberFormatCategory: 2, numberFormatChoices: 1, errorText: 2,
      scanReferences: 2, validateSheetName: 2, sameValue: 2, canonValue: 1, dateToSerial: 2, serialToDate: 2, debugGeneratedCode: 2, internStyle: 2, columnWidthFromPoints: 1, columnWidthToPoints: 1, escapeColumn: 1, formatClipboard: 2, parseClipboard: 1,
      parseCell: 1, cellAddress: 2, colLetters: 1, colNumber: 1, parseRange: 1, rangeAddress: 1, parseLines: 2, linesKey: 3, patchLines: 3, rectsMeet: 2, rectContains: 2, quoteSheetIfNeeded: 2, nameKey: 1, sameName: 2, outlineGroups: 1,
    }],
    ['xlcalc/a1', a1, { parseCell: 1, cellAddress: 2, colLetters: 1, colNumber: 1, parseRange: 1, rangeAddress: 1, parseLines: 2, linesKey: 3, patchLines: 3, rectsMeet: 2, rectContains: 2, quoteSheetIfNeeded: 2, nameKey: 1, sameName: 2, outlineGroups: 1 }],
    ['xlcalc/text', text, { scanReferences: 2, quoteSheetIfNeeded: 2, columnWidthFromPoints: 1, columnWidthToPoints: 1, escapeColumn: 1, formatClipboard: 2, parseClipboard: 1, numberFormatCategory: 2, numberFormatChoices: 1 }],
    ['xlcalc/xlsx', xlsx, { readXlsx: 2, writeXlsx: 2, fromLegacy: 2 }],
    ['xlcalc/worker', worker, { createCalcPool: 1 }],
  ]

  for (const [entry, mod, table] of FUNCTIONS)
    it(`${entry}: each function by its parameters, refusing an options object past them`, async () => {
      const fns = Object.fromEntries(Object.entries(mod).filter(([k, f]) => typeof f === 'function' && !/Error$/.test(k)))
      expect(Object.fromEntries(Object.entries(fns).map(([k, f]) => [k, (f as () => void).length]))).toEqual(table)
      for (const [name, max] of Object.entries(table)) {
        const f = fns[name] as (...a: unknown[]) => unknown
        const refused = `takes ${max} argument${max === 1 ? '' : 's'}, got an object as argument ${max + 1} (table is an option of toLocal and toInvariant)`
        if (entry === 'xlcalc/xlsx' && name !== 'fromLegacy') await expect(f(...Array(max).fill(undefined), { table: 'T' }), name).rejects.toThrow(new InputError(name, refused))
        else expect(() => f(...Array(max).fill(undefined), { table: 'T' }), name).toThrow(new InputError(name, refused))
      }
    })

  it("toLocal, toInvariant and fromLegacy with an unknown option; a callback's index and array pass", () => {
    expect(() => toInvariant('СУММ(A1)', ruRU, { tabel: 'T' } as never)).toThrow('opts: unknown field "tabel" (expected table, externalBooks)')
    expect(() => (toInvariant as (...a: unknown[]) => unknown)('СУММ(A1)', ruRU, { table: 'T' }, { table: 'T' })).toThrow('toInvariant: takes 3 arguments, got an object as argument 4 (table is an option of toLocal and toInvariant)')
    expect(() => toLocal('SUM(A1)', ruRU, { tabel: 'T' } as never)).toThrow('opts: unknown field "tabel" (expected table, externalBooks)')
    expect(() => toLocal('SUM(A1)', ruRU, 'T' as never)).toThrow(InputError)
    expect(toLocal('T[[#This Row],[Qty]]*2', ruRU, { table: 'T' })).toEqual({ ok: true, formula: '[@Qty]*2' })
    expect(() => xlsx.fromLegacy('A1:A2*2', { sheets: 'S' } as never)).toThrow('ctx: unknown field "sheets" (expected names, sheet)')
    expect(xlsx.fromLegacy('A1:A2*2', { sheet: 'S' })).toBe('@A1:A2*2')
    expect(['A1', 'B2'].map(parseCell)).toEqual([{ row: 1, col: 1 }, { row: 2, col: 2 }])
    expect([1, 2].map(root.colLetters)).toEqual(['A', 'B'])
    expect(() => compute(book(), { clock: CLOCK, formulaLocal: true } as never)).toThrow('opts: unknown field "formulaLocal" (expected clock, externals, execution, display, limits, cached) (formulaLocal is an option of getViewport)')
  })
})
