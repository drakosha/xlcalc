/**
 * The old names of the API are gone, so that a host still on one fails where it calls it, not later: a function or a
 * method by its old name is no function (a TypeError at the call, and no type to compile against) in a synchronous
 * session, in the same thread and in a thread of the pool, and no entry of the package exports one. Input that gives a
 * number format by its old name `format` is an `InputError` that names `numberFormat`, in every input that takes a
 * number format, `viewportOf`'s saved book too (`viewport-saved.test.ts`); `writeXlsx` refuses a book with its result in
 * the options.
 */
import { afterAll, describe, expect, it } from 'vitest'
import * as root from '../src/index.js'
import { compute, createInlineCalc, createSession, displayValue, entryWidth, InputError, parseEntry, type AsyncSession, type BookInput, type Calc, type Session } from '../src/index.js'
import * as a1 from '../src/a1.js'
import * as contract from '../src/contract.js'
import * as text from '../src/text.js'
import * as typedEntry from '../src/typed-entry.js'
import * as worker from '../src/worker/index.js'
import * as xlsx from '../src/xlsx.js'
import { CLOCK } from './async-books.js'
import { closeCalcs, keptPool, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const book = (): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { formula: 'A1*10' } } }] })

/** A session's methods by their old names, each with the name it has now. */
const METHODS = { book: 'getBook', layout: 'getLayout', extent: 'getUsedRange', cells: 'getCells', validation: 'getValidation', result: 'getResult', hash: 'getHash', restructure: 'editStructure' }

/** The functions and objects the root exported by their old names. */
const ROOT = ['openSession', 'restructure', 'engine', 'fromLegacy', 'liveWorkbook']

/** The hint of a refusal of the old field. */
const HINT = '(format is the old name of numberFormat)'

// the types by their old names are gone too: each line fails to compile if one comes back (`npm run typecheck`)
// @ts-expect-error the name is StructureResult
type R1 = import('../src/index.js').RestructureResult
// @ts-expect-error the name is SessionStructureResult
type R2 = import('../src/index.js').SessionRestructureResult
// @ts-expect-error the name is AsyncStructureResult
type R3 = import('../src/index.js').AsyncRestructureResult
// @ts-expect-error the name is ConditionalAppearance
type R4 = import('../src/index.js').ConditionalLook
// @ts-expect-error the engine object is gone, and its type
type R5 = import('../src/index.js').FormulaEngine
// @ts-expect-error LegacyContext is xlcalc/xlsx's
type R6 = import('../src/index.js').LegacyContext
// @ts-expect-error the name is LiveBook
type R7 = import('../src/index.js').LiveWorkbook

describe('no entry of the package exports an old name', () => {
  it('the root has none of them; fromLegacy is xlcalc/xlsx\'s alone', () => {
    expect(ROOT.filter((name) => name in root)).toEqual([])
    expect(typeof xlsx.fromLegacy).toBe('function')
  })

  it.each([['xlcalc/contract', contract], ['xlcalc/a1', a1], ['xlcalc/text', text], ['xlcalc/entry', typedEntry], ['xlcalc/xlsx', xlsx], ['xlcalc/worker', worker]] as const)('%s has none of them', (_, mod) => {
    expect([...ROOT.filter((name) => name !== 'fromLegacy'), ...Object.keys(METHODS)].filter((name) => name in mod)).toEqual([])
  })
})

describe('a call by an old name fails at once', () => {
  /** Each kind of session as it is given out, not wrapped: a synchronous one, the same thread's, a thread's. */
  const SESSIONS: [string, () => Promise<{ session: Session | AsyncSession; calc?: Calc }>][] = [
    ['a synchronous session', async () => ({ session: createSession(book(), { clock: CLOCK }) })],
    ['the same thread', async () => {
      const calc = createInlineCalc()
      return { session: await calc.createSession(book(), { clock: CLOCK }), calc }
    }],
    ['a thread', async () => {
      const calc = await keptPool()
      return { session: await calc.createSession(book(), { clock: CLOCK }), calc }
    }],
  ]

  it.each(SESSIONS)('%s: each method by its old name is no function, its new name answers', async (kind, open) => {
    const { session, calc } = await open()
    const methods = session as unknown as Record<string, ((...a: unknown[]) => unknown) | undefined>
    for (const [old, now] of Object.entries(METHODS)) {
      expect(methods[old], old).toBeUndefined()
      expect(typeof methods[now], now).toBe('function')
      expect(() => methods[old]!('S'), old).toThrow(TypeError)
    }
    expect(await session.getHash()).toBe(compute(book(), { clock: CLOCK }).hash)
    if (calc) {
      const c = calc as unknown as Record<string, unknown>
      expect(c.openSession, 'openSession').toBeUndefined()
      expect(() => (c.openSession as (b: BookInput) => unknown)(book())).toThrow(TypeError)
    }
    await session.close()
    if (kind === 'the same thread') await calc!.close()
  })

  it('the root: openSession, restructure and engine are no functions', () => {
    const r = root as unknown as Record<string, ((...a: unknown[]) => unknown) | undefined>
    expect(() => r.openSession!(book())).toThrow(TypeError)
    expect(() => r.restructure!(book(), { op: 'insertRows', sheet: 'S', row: 1 })).toThrow(TypeError)
    expect(() => (r.engine as unknown as { compute: () => unknown }).compute()).toThrow(TypeError)
  })

  it('writeXlsx refuses a book with its result in the options: the book and its result go as one snapshot', async () => {
    const b = book()
    await expect(xlsx.writeXlsx(b as never, { result: compute(b, { clock: CLOCK }) } as never)).rejects.toThrow(new InputError('snapshot', 'unknown field "sheets" (expected book, result) (a book is not a snapshot: pass { book, result })'))
  })
})

describe('the number format by its old name is refused, naming numberFormat', () => {
  /** A book with one place that gives the old name. */
  const sheet = (x: object) => ({ sheets: [{ name: 'S', cells: { A1: { value: 1 } }, ...x }] }) as BookInput

  it('in a book: a cell, rows, columns, a conditional format\'s style', () => {
    expect(() => compute(sheet({ cells: { A1: { value: 1, format: '0.00' } } }))).toThrow(new InputError('book.sheets[0].cells["A1"]', `unknown field "format" (expected value, formula, numberFormat, style, meta, array, note, link, held) ${HINT}`))
    expect(() => compute(sheet({ rows: { 2: { format: '0.00' } } }))).toThrow(new InputError('book.sheets[0].rows["2"]', `unknown field "format" (expected height, bestFit, style, numberFormat, hidden, outlineLevel, collapsed) ${HINT}`))
    expect(() => compute(sheet({ columns: { B: { format: '0.00' } } }))).toThrow(new InputError('book.sheets[0].columns["B"]', `unknown field "format" (expected width, bestFit, style, numberFormat, hidden, outlineLevel, collapsed) ${HINT}`))
    expect(() => compute(sheet({ conditionalFormats: [{ type: 'expression', ranges: ['A1'], formula: 'TRUE', style: { format: '0.0' } }] }))).toThrow(new InputError('book.sheets[0].conditionalFormats[0].style', `unknown field "format" (expected font, fill, border, numberFormat) ${HINT}`))
  })

  it('displayValue, entryWidth and parseEntry\'s cell', () => {
    expect(() => displayValue(0.25, { format: '0.0%' } as never)).toThrow(new InputError('opts', `unknown field "format" (expected numberFormat, locale, width, dateSystem) ${HINT}`))
    expect(() => entryWidth(46096, { format: 14 } as never)).toThrow(HINT)
    expect(() => parseEntry('5', { cell: { format: '0%' } as never })).toThrow(new InputError('opts.cell', `unknown field "format" (expected numberFormat, style, meta, note, link) ${HINT}`))
    // a cell of a window has more fields than these, which are not read
    expect(parseEntry('5', { cell: { numberFormat: '0%', text: '1%' } as never })).toEqual({ ok: true, input: { value: 0.05, numberFormat: '0%' } })
  })

  describe.each(KINDS)('%s', (kind) => {
    it('a book, setCells (null too), replaceCells, patchCells, setLayout of lines and of conditional formats, validateEntry', async () => {
      const open = await opener(kind)
      await expect(open(sheet({ cells: { A1: { value: 1, format: '0.00' } } }), { clock: CLOCK })).rejects.toThrow(HINT)
      const s = await open(book(), { clock: CLOCK })
      const refused = async (call: () => Promise<unknown>, path: string) => {
        const e = await call().then(() => null, (x: unknown) => x)
        expect(e, path).toBeInstanceOf(InputError)
        expect((e as InputError).message, path).toMatch(new RegExp(`^${path.replace(/[[\]."]/g, '\\$&')}: unknown field "format" \\(expected [^)]*\\) \\(format is the old name of numberFormat\\)$`))
      }
      await refused(() => s.setCells([{ sheet: 'S', cell: 'B1', input: { value: 0.5, format: '0.0%' } as never }]), 'edits[0].input')
      await refused(() => s.setCells([{ sheet: 'S', cell: 'B1', input: { format: null } as never }]), 'edits[0].input')
      await refused(() => s.replaceCells([{ sheet: 'S', cell: 'B1', input: { value: 3, format: '0.0' } as never }]), 'edits[0].input')
      await refused(() => s.patchCells([{ sheet: 'S', cell: 'B1', set: { format: '0.000' } as never }]), 'patches[0].set')
      await refused(() => s.setLayout([{ sheet: 'S', rows: '5', set: { format: '0.0000' } as never }]), 'edits[0].set')
      await refused(() => s.setLayout([{ sheet: 'S', conditionalFormats: [{ type: 'expression', ranges: ['E1'], formula: 'TRUE', style: { format: '0' } as never }] }]), 'edits[0].conditionalFormats[0].style')
      await refused(() => s.validateEntry('S', 'A1', { value: 2, format: '0' } as never), 'input')
      // nothing changed: the book is as it was opened
      expect((await s.getBook()).sheets[0].cells).toEqual(book().sheets[0].cells)
      expect(s.revision).toBe(0)
      await s.close()
    })
  })
})
