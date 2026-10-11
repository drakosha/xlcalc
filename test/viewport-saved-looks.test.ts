/**
 * `viewportOf` with `appearance`: a saved version's conditional formats computed over its result's values, its formulas
 * not computed, against the window a live session over the book gives on the result's clock, in the three kinds of
 * session. The cases of the Excel sets of conditional formats (`bench/coverage/cond/`: every kind of rule, priorities,
 * `stopIfTrue`, a rule's number format, ru-RU, books after structure operations), and a book whose rules of every kind
 * read formulas (of other cells, of another sheet, through a name, a spill, the cells a large spill's result leaves
 * out, TODAY), in en-US and ru-RU, with the display and without. TODAY in a rule is the result's day; the state the
 * rules are computed over is kept with the objects; where a saved window differs from a live one, pinned.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, createSession, editStructure, viewportOf, type BookInput, type ConditionalFormat, type Execution, type Viewport, type ViewportOptions } from '../src/index.js'
import type { CondSet } from '../bench/coverage/cond/cases.js'
import { takenBook } from '../bench/coverage/stored.js'
import { CLOCK } from './async-books.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'
import './locales.js'

afterAll(closeCalcs)

/** The sets of `bench/coverage/cond/`. */
const SETS = ['looks', 'looks-ru', 'files', 'struct-held', 'held', 'held-2', 'both', 'made', 'edge-held', 'formats-ru-held']

const ALL: ViewportOptions = { formulaLocal: true, diagnostics: true, editText: true }
const RED = { fill: { color: '#FFC7CE' } }

/** A live window as a saved one with `appearance` gives it: without the session's revision. */
const live = ({ revision: _r, ...rest }: Viewport) => rest

/** A book and its result as a host saves a version. */
const saved = (book: BookInput, o: { execution?: Execution; display?: boolean } = {}) => JSON.parse(JSON.stringify({ book, result: compute(book, { clock: CLOCK, ...o }) })) as { book: BookInput; result: ReturnType<typeof compute> }

/** The cells of a window that have an appearance. */
const painted = (v: { cells: Record<string, { appearance?: unknown }> }) => Object.keys(v.cells).filter((a) => v.cells[a].appearance)

describe.each(KINDS)('the Excel sets of conditional formats, %s', (kind) => {
  it.each(SETS)('%s: each case read as a live session reads it', async (set) => {
    const open = await opener(kind)
    const mod = (await import(`../bench/coverage/cond/${set}.ts`)) as CondSet
    const built = mod.buildCondCases()
    let looks = 0
    for (const c of built.cases) {
      const taken = takenBook(c.make && mod.expectedBook ? mod.expectedBook(c) : built.books[c.book])
      const r = c.op ? editStructure(taken, c.op) : { ok: true as const, book: taken }
      if (!r.ok) continue
      const version = saved(r.book)
      const s = await open(r.book, { clock: CLOCK })
      for (const [sheet, range] of c.read) {
        const v = viewportOf(version, sheet, range, { ...ALL, appearance: true })
        expect(v, `${c.id}: ${sheet}!${range}`).toEqual(live(await s.getViewport(sheet, range, ALL)))
        looks += painted(v).length
      }
      await s.close()
    }
    expect(looks).toBeGreaterThan(0)
  })
})

/** A rule over `ranges` with a red fill but for its own fields. */
const red = (rule: Record<string, unknown>, ...ranges: string[]) => ({ ranges, style: RED, ...rule }) as ConditionalFormat

/**
 * Rules of every kind over formulas: of other cells, of another sheet, through a name, over a spill (`E1#`), over the
 * cells of a large spill that its result leaves out (M70000), dates by TODAY, a rule's number format, empty cells.
 */
function rulesBook(locale?: string): BookInput {
  const rows = (col: string, formula: (r: number) => string, n: number) => Object.fromEntries(Array.from({ length: n }, (_, i) => [`${col}${i + 1}`, { formula: formula(i + 1) }]))
  const cells: BookInput['sheets'][number]['cells'] = {
    A1: { value: 4 },
    A2: { formula: 'A1*2' },
    A3: { formula: "'T'!A1+1" },
    A4: { formula: 'Lim*2' },
    A5: { formula: 'NA()' },
    A6: { value: 'x' },
    A7: { formula: 'SUM(E1#)' },
    A8: { formula: 'A2-A1' },
    A9: { value: -2.5 },
    B1: { formula: 'A1+1' },
    C1: { formula: '"ab"&"c"' },
    C2: { value: 'xyz' },
    C3: { formula: 'UPPER(C2)' },
    C4: { formula: 'A9*1' },
    D1: { formula: 'TODAY()' },
    D2: { formula: 'TODAY()-1' },
    D3: { formula: 'TODAY()+1' },
    D4: { formula: 'EDATE(TODAY(),-1)' },
    D5: { formula: 'TODAY()-6' },
    D6: { value: { date: '2026-01-01' } },
    E1: { formula: 'SEQUENCE(3)' },
    ...rows('G', (r) => `ROW()^2`, 8),
    ...rows('H', (r) => `ROW()-4`, 8),
    ...rows('I', (r) => `ROW()*10`, 8),
    J1: { formula: '1/4' },
    J2: { formula: 'J1*3' },
    K1: { formula: 'A1' },
    K2: { formula: 'A8' },
    K3: { value: 9 },
    L2: { formula: '""' },
    M1: { formula: 'N1:N70000*2' },
    N1: { value: 1 },
  }
  const conditionalFormats: ConditionalFormat[] = [
    red({ type: 'cellIs', operator: 'greaterThan', formulas: ['$B$1'] }, 'A1:A9'),
    { type: 'expression', ranges: ['A1:A9'], formula: 'A1>=Lim', style: { font: { bold: true } }, stopIfTrue: true },
    { type: 'top10', ranges: ['A1:A9'], rank: 2, style: { font: { italic: true } } },
    { type: 'aboveAverage', ranges: ['A1:A9'], below: true, style: { font: { color: '#9C0006' } } },
    { type: 'containsErrors', ranges: ['A1:A9'], style: { fill: { color: '#000000' } } },
    red({ type: 'containsText', text: 'B' }, 'C1:C4'),
    red({ type: 'containsText', text: ',' }, 'C1:C4'),
    red({ type: 'timePeriod', period: 'today' }, 'D1:D6'),
    { type: 'timePeriod', ranges: ['D1:D6'], period: 'yesterday', style: { font: { bold: true } } },
    { type: 'timePeriod', ranges: ['D1:D6'], period: 'last7Days', style: { font: { italic: true } } },
    { type: 'timePeriod', ranges: ['D1:D6'], period: 'lastMonth', style: { fill: { color: '#00FF00' } } },
    red({ type: 'expression', formula: 'SUM($E$1#)=6' }, 'F1:F3'),
    { type: 'colorScale', ranges: ['G1:G8'], points: [{ value: { type: 'min' }, color: '#F8696B' }, { value: { type: 'percentile', value: '50' }, color: '#FFEB84' }, { value: { type: 'formula', value: '$B$1*10' }, color: '#63BE7B' }] },
    { type: 'dataBar', ranges: ['H1:H8'], min: { type: 'autoMin' }, max: { type: 'autoMax' }, color: '#638EC6', negativeColor: '#FF0000' },
    { type: 'iconSet', ranges: ['I1:I8'], icons: '3Arrows', points: [{ value: { type: 'percent', value: '0' } }, { value: { type: 'percent', value: '33' } }, { value: { type: 'percent', value: '67' } }] },
    { type: 'cellIs', ranges: ['J1:J2'], operator: 'lessThan', formulas: ['1'], style: { numberFormat: '0.00%', font: { color: '#0000FF' } } },
    red({ type: 'duplicateValues' }, 'K1:K3'),
    { type: 'uniqueValues', ranges: ['K1:K3'], style: { font: { bold: true } } },
    red({ type: 'containsBlanks' }, 'L1:L4'),
    red({ type: 'cellIs', operator: 'equal', formulas: ['0'] }, 'M1:M5', 'M69998:M70000'),
    red({ type: 'expression', formula: 'ISNUMBER($M$70000)' }, 'O1:O2'),
  ]
  return {
    ...(locale ? { locale } : {}),
    names: [{ name: 'Lim', ref: "'T'!$A$2" }],
    sheets: [
      { name: 'S', cells, conditionalFormats },
      { name: 'T', cells: { A1: { formula: '2*3' }, A2: { value: 5 } }, conditionalFormats: [red({ type: 'expression', formula: "A1>'S'!$A$1" }, 'A1:A2')] },
    ],
  }
}

const WINDOWS: [string, string][] = [
  ['S', 'A1:O9'],
  ['S', 'M69995:O70002'],
  ['T', 'A1:B3'],
]

describe.each(['closures', 'eval'] as const)('rules of every kind over formulas, execution %s', (execution: Execution) => {
  describe.each(KINDS)('%s', (kind) => {
    it.each([['en-US', undefined], ['ru-RU', 'ru-RU']] as const)('%s, with the display and without: the live window', async (_, locale) => {
      const open = await opener(kind)
      for (const display of [true, false]) {
        const book = rulesBook(locale)
        const version = saved(book, { execution, display })
        const s = await open(book, { clock: CLOCK, execution, display })
        for (const [sheet, range] of WINDOWS) expect(viewportOf(version, sheet, range, { ...ALL, appearance: true }), `${sheet}!${range}${display ? '' : ' without the display'}`).toEqual(live(await s.getViewport(sheet, range, ALL)))
        await s.close()
      }
    })
  })
})

describe('the appearance of a saved version', () => {
  it('each rule gives what it gives a live session: other cells, another sheet, a name, a spill, a large spill, TODAY, a number format', () => {
    const v = viewportOf(saved(rulesBook()), 'S', 'A1:O9', { appearance: true })
    // A2 (8) over B1 (5) and the name's 5, A3 reads T!A1 (7), A7 sums the spill E1# (6); the bold rule stops the rules after it
    expect(v.cells.A2.appearance).toEqual({ style: { fill: RED.fill, font: { bold: true } } })
    expect(v.cells.A7.appearance).toEqual({ style: { fill: RED.fill, font: { bold: true } } })
    expect(v.cells.A5.appearance).toEqual({ style: { fill: { color: '#000000' } } })
    // the day of the result's clock (2026-10-04): D1 today, D2 yesterday, D5 six days back (in September), D4 a month back
    const green = { color: '#00FF00' }
    expect([v.cells.D1, v.cells.D2, v.cells.D5, v.cells.D4, v.cells.D6].map((c) => c.appearance)).toEqual([{ style: { fill: RED.fill, font: { italic: true } } }, { style: { font: { bold: true, italic: true } } }, { style: { font: { italic: true }, fill: green } }, { style: { fill: green } }, undefined])
    expect(painted(v).filter((a) => /^[FO]/.test(a))).toEqual(['F1', 'O1', 'F2', 'O2', 'F3'])
    expect(v.cells.J1).toEqual({ value: 0.25, text: '25.00%', formula: '1/4', appearance: { style: { numberFormat: '0.00%', font: { color: '#0000FF' } } } })
    expect([v.cells.H1.appearance?.bar?.color, v.cells.I8.appearance?.icon, v.cells.G8.appearance?.style?.fill?.color]).toEqual(['#FF0000', { set: '3Arrows', index: 2 }, '#63BE7B'])
    // an empty cell with an appearance is listed, as a session lists it
    expect([v.cells.L1, v.cells.L2.appearance]).toEqual([{ value: null, text: '', appearance: { style: RED } }, { style: RED }])
    expect(viewportOf(saved(rulesBook()), 'S', 'M69999:M70001', { appearance: true }).cells.M70000).toEqual({ value: 0, text: '0', spillOf: 'M1', appearance: { style: RED } })
  })

  it("TODAY in a rule is the result's day, not the day a window is read", () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: { date: '2026-10-04' } } }, conditionalFormats: [red({ type: 'timePeriod', period: 'today' }, 'A1')] }] }
    const v = viewportOf(saved(book), 'S', 'A1', { appearance: true })
    expect(v.cells.A1.appearance).toEqual({ style: RED })
    expect(createSession(book, { clock: { now: '2026-10-05T12:00:00Z', seed: 7 } }).getViewport('S', 'A1').cells.A1.appearance).toBeUndefined()
  })

  it('without the option the window has no appearance, the text by the cell’s own format', () => {
    const version = saved(rulesBook())
    const v = viewportOf(version, 'S', 'J1:L4')
    expect(painted(v)).toEqual([])
    expect(v.cells.J1).toEqual({ value: 0.25, text: '0.25', formula: '1/4' })
    expect(v.cells.L1).toBeUndefined()
  })

  it('the state the rules are computed over is built once for the objects: the later windows read it, not the result again', () => {
    const version = saved(rulesBook())
    expect(viewportOf(version, 'S', 'A2', { appearance: true }).cells.A2.appearance).toBeDefined()
    // a host changes neither object; a change after the first window does not reach the rules
    version.result.sheets[0].cells.A2.value = 0
    expect(viewportOf(version, 'S', 'A2', { appearance: true }).cells.A2).toMatchObject({ value: 0, appearance: { style: { fill: RED.fill } } })
    expect(viewportOf({ ...version }, 'S', 'A2', { appearance: true }).cells.A2.appearance).toBeDefined()
  })

  it('where a saved window differs from a live one: numbers past 15 digits, SUBTOTAL over SUBTOTAL, external functions', () => {
    const look = (book: BookInput, range: string, opts = {}) => {
      const v = viewportOf(saved(book), 'S', range, { appearance: true })
      const l = createSession(book, { clock: CLOCK, ...opts }).getViewport('S', range)
      return [painted(v), painted(l)]
    }
    // the result holds 0.1+0.2 as 0.3, to 15 digits: a duplicate of 0.3, which the session's full number is not
    expect(look({ sheets: [{ name: 'S', cells: { A1: { formula: '0.1+0.2' }, A2: { value: 0.3 } }, conditionalFormats: [red({ type: 'duplicateValues' }, 'A1:A2')] }] }, 'A1:A2')).toEqual([['A1', 'A2'], []])
    // a value holds no formula: SUBTOTAL in a rule counts a cell with a SUBTOTAL of its own, which the session leaves out
    const nested: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1 }, A2: { value: 2 }, A3: { formula: 'SUBTOTAL(9,A1:A2)' } }, conditionalFormats: [red({ type: 'expression', formula: 'SUBTOTAL(9,$A$1:$A$3)=3' }, 'B1')] }] }
    expect(look(nested, 'A1:B3')).toEqual([[], ['B1']])
    // a result holds no externals: a rule's external function does not read
    const ext: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1 } }, conditionalFormats: [red({ type: 'expression', formula: 'LIMIT()<5' }, 'A1')] }] }
    expect(look(ext, 'A1', { externals: { functions: ['LIMIT'], values: { 'LIMIT()': 3 } } })).toEqual([[], ['A1']])
  })
})
