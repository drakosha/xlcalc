/**
 * The values hash against the canonical text built here from its rules and hashed by node:crypto, independently of
 * `src/canon.ts`: `Session.hash()`, `result().hash`, `compute().hash` and `valuesHash` agree with it on edge books
 * (sheet and address order, rounding, non-finite numbers, Unicode, ru-RU, 1904, spills and their fills) and seeded
 * random books, after cell edits and structure operations, in a synchronous session, in the same thread and in a thread,
 * with closures and eval.
 */
import { createHash } from 'node:crypto'
import { afterAll, describe, expect, it } from 'vitest'
import { valuesHash, type FillLine } from '../src/canon.js'
import { CONFIG_VERSION } from '../src/config.js'
import { compute, createInlineCalc, openSession, type BookInput, type CellEdit, type CellInput, type CellValue, type ComputeResult, type Execution, type Session, type StructureOp } from '../src/index.js'
import { clockSerial, createState, formatAddress, outValue, toCellValue } from '../src/internal.js'
import { mulberry32, synthBook } from '../src/synth.js'
import { pool } from './fixtures/pool.js'

const CLOCK = { now: '2026-10-04T12:00:00Z', seed: 7 }

// ---------- the reference: the canonical text by its rules ----------

const byCode = (a: string, b: string) => (a < b ? -1 : a > b ? 1 : 0)

/** Row and column of an address, for the order by row, then column. */
function position(addr: string): [number, number] {
  const m = /^([A-Z]+)([0-9]+)$/.exec(addr)!
  let col = 0
  for (const ch of m[1]) col = col * 26 + ch.charCodeAt(0) - 64
  return [Number(m[2]), col]
}

const byAddress = (a: string, b: string) => {
  const [x, y] = [position(a), position(b)]
  return x[0] - y[0] || x[1] - y[1]
}

/** A value as a line writes it; a result's numbers are rounded already, a non-finite one is #NUM!, an empty fill `_`. */
function valueText(v: CellValue): string {
  if (v === null) return '_'
  if (typeof v === 'number') return Number.isFinite(v) ? `n ${String(v)}` : 'e NUM'
  if (typeof v === 'string') return `s ${JSON.stringify(v)}`
  if (typeof v === 'boolean') return v ? 'b TRUE' : 'b FALSE'
  return `e ${v.error}`
}

/** The unstored cells of a book's large spills, from a state computed afresh, one line per rectangle. */
function fillsOf(book: BookInput, execution: Execution): FillLine[] {
  const state = createState(book, { now: clockSerial(CLOCK), seed: CLOCK.seed, execution })
  const out: FillLine[] = []
  for (const x of state.spills.on.values()) for (const f of x.fills) out.push({ sheet: state.sheetNames[x.s], range: `${formatAddress(f.r1, f.c1)}:${formatAddress(f.r2, f.c2)}`, value: outValue(toCellValue(x.fills[0].v)) })
  return out
}

/** The header, then the non-empty cells by sheet name and address, then the fills by sheet and top-left cell, joined by `\n`. */
function referenceText(r: ComputeResult, book: BookInput, fills: FillLine[]): string {
  const head = `${CONFIG_VERSION}${book.locale !== undefined && book.locale !== 'en-US' ? `;locale=${book.locale}` : ''}${book.dateSystem === 1904 ? ';date1904' : ''}\n`
  const lines: string[] = []
  for (const sheet of [...r.sheets].sort((a, b) => byCode(a.name, b.name)))
    for (const addr of Object.keys(sheet.cells).sort(byAddress)) {
      const v = sheet.cells[addr].value
      if (v !== null) lines.push(`${JSON.stringify(sheet.name)}!${addr} ${valueText(v)}`)
    }
  const top = (f: FillLine) => f.range.split(':')[0]
  for (const f of [...fills].sort((a, b) => byCode(a.sheet, b.sheet) || byAddress(top(a), top(b)))) lines.push(`${JSON.stringify(f.sheet)}!${f.range} ${valueText(f.value)}`)
  return head + lines.join('\n')
}

const sha256 = (text: string) => createHash('sha256').update(text, 'utf8').digest('hex')

/**
 * A synchronous session's `hash()` and `result().hash` (with and without display), `compute()` of its book and
 * `valuesHash` over the computed values (prepared and not) all equal the reference hash; returns it.
 */
function agree(s: Session, execution: Execution, label: string): string {
  const h = s.hash()
  const book = s.book()
  const fresh = compute(book, { clock: CLOCK, execution })
  const fills = fillsOf(book, execution)
  const expected = sha256(referenceText(fresh, book, fills))
  expect([h, s.result().hash, fresh.hash, compute(book, { clock: CLOCK, execution, display: false }).hash], label).toEqual([expected, expected, expected, expected])
  const values = new Map(fresh.sheets.map((x) => [x.name, new Map(Object.entries(x.cells).map(([a, c]): [string, CellValue] => [a, c.value]))]))
  expect([valuesHash(values, CONFIG_VERSION, fills, book.locale, book.dateSystem, true), valuesHash(values, CONFIG_VERSION, fills, book.locale, book.dateSystem)], label).toEqual([expected, expected])
  return expected
}

// ---------- edge books ----------

const sheet = (name: string, cells: Record<string, CellInput>) => ({ name, cells })

/** Names out of their order by code unit (upper case before lower, Latin before Cyrillic, Ё before А), with quotes. */
const NAMES = ['b', 'C', 'Альфа', 'ёлка', 'Ёж', 'A "q" B', "It's", 'Z1']
/** Addresses out of their order by text: A9/A10, Z1/AA1, the last row and column. */
const ADDRS = ['A10', 'A9', 'AA1', 'Z1', 'B2', 'A1', 'XFD1048576', 'A1048576', 'XFD1', 'AB10', 'AZ9', 'BA9']

const VALUES: (CellInput['value'])[] = [
  -0, 0.1 + 0.2, 999999999999999.6, 1.2345678901234567e-20, 123456789012345678, 1e21, 5e-324, -1.7976931348623155e308, Number.MAX_VALUE,
  '', 'текст\n"', '🐻 \\ \t   \r\n', '\ud800 lone', 'TRUE', '1.5', true, false, null, { date: '2026-10-04' },
]

const FORMULAS = ['1/0', 'NA()', 'SQRT(-1)', '0.1+0.2', '1E308*10', '"ы"&"é"', 'TRUE', '1/3', 'RAND()', 'NOW()', 'Q1*1', 'Q1+Q1', 'X1', 'X2']

const EDGE: [string, BookInput][] = [
  ['sheet and address order', { sheets: NAMES.map((n, i) => sheet(n, Object.fromEntries(ADDRS.map((a, j) => [a, { value: i * 100 + j }])))) }],
  [
    'values',
    {
      sheets: [
        sheet('V', {
          ...Object.fromEntries(VALUES.map((v, i) => [`A${i + 1}`, { value: v }])),
          ...Object.fromEntries(FORMULAS.map((f, i) => [`C${i + 1}`, { formula: f }])),
          Q1: { value: Number.MAX_VALUE },
          X1: { formula: 'X2' },
          X2: { formula: 'X1' },
          D1: { format: '0.00' },
        }),
      ],
    },
  ],
  [
    'spills over cells with only a format or a style',
    {
      styles: [{}],
      sheets: [
        sheet('S', { A1: { formula: 'SEQUENCE(3)' }, A2: { format: '0.00' }, A3: { style: 0 }, B1: { formula: 'SEQUENCE(2,2)' }, C2: { style: 0, format: '0%' }, E1: { formula: 'SEQUENCE(3)' }, E2: { value: 1 }, G1: { formula: 'A1#*2' } }),
      ],
    },
  ],
  [
    'large spills and their fills',
    {
      styles: [{}],
      sheets: [
        sheet('T', { B1: { value: 1 }, B2: { value: 2 }, B3: { value: 3 }, C1: { formula: 'B:B*2' }, C7: { format: '0.00' }, E1: { formula: 'SEQUENCE(100000,1,0,0)' }, E5: { style: 0 }, G1: { formula: 'IF(SEQUENCE(300,300)>0,0)' } }),
        sheet('Fill', { A1: { formula: 'SEQUENCE(100000,2,0,0)' }, D1: { formula: 'IF(SEQUENCE(100000)>0,"ы")' } }),
        sheet('a', { A1: { formula: 'IF(SEQUENCE(100000)>0,TRUE)' } }),
      ],
    },
  ],
  ['empty sheets and empty cells', { sheets: [sheet('E', {}), sheet('D', { A1: { value: null }, B2: { format: '0' } })] }],
]

const withSettings = (book: BookInput, settings: Partial<BookInput>): BookInput => ({ ...book, ...settings })
EDGE.push(['values in ru-RU and the 1904 system', withSettings(EDGE[1][1], { locale: 'ru-RU', dateSystem: 1904 })], ['large spills in ru-RU', withSettings(EDGE[3][1], { locale: 'ru-RU' })])

// ---------- seeded random books, edits and operations ----------

const pick = <T>(rnd: () => number, xs: readonly T[]): T => xs[Math.floor(rnd() * xs.length)]
const COLS = ['A', 'B', 'C', 'Z', 'AA', 'AB', 'AZ', 'BA']
const ROWS = [1, 2, 3, 9, 10, 11, 99, 100]

/** A random cell input; `styled`: the book has a style to refer to. */
function randomCell(rnd: () => number, names: readonly string[], styled: boolean): CellInput {
  const x = rnd()
  if (x < 0.4) return { value: pick(rnd, VALUES) }
  if (x < 0.5) return { value: Math.round((rnd() - 0.5) * 1e6) / 1e3 }
  if (x < 0.6) return styled ? pick(rnd, [{ format: '0.00' }, { style: 0 }, { value: null, style: 0 }]) : { format: '0.00' }
  if (x < 0.65) return { formula: pick(rnd, ['SEQUENCE(3,2)', 'SEQUENCE(100000,1,0,0)', 'B:B*2', 'IF(SEQUENCE(40,40)>0,"x")']) }
  // N() and not arithmetic over the reference: 0b33695 throws on arithmetic over text with a line break ("a\nb"*1)
  if (x < 0.85) {
    const ref = `${pick(rnd, ['', `'${pick(rnd, names)}'!`])}${pick(rnd, COLS)}${pick(rnd, ROWS)}`
    return { formula: pick(rnd, [ref, `N(${ref})*${pick(rnd, ['2', '0.1', '-1'])}`, `LEN(${ref})`]) }
  }
  return { formula: pick(rnd, FORMULAS) }
}

/** A book of 1 to 4 sheets with names from NAMES, cells at addresses around the order's boundaries. */
function randomBook(seed: number): BookInput {
  const rnd = mulberry32(seed)
  const names = [...NAMES].sort(() => rnd() - 0.5).slice(0, 1 + Math.floor(rnd() * 4))
  return {
    styles: [{}],
    ...(rnd() < 0.3 ? { locale: 'ru-RU' } : {}),
    ...(rnd() < 0.2 ? { dateSystem: 1904 as const } : {}),
    sheets: names.map((name) => {
      const cells: Record<string, CellInput> = {}
      for (let i = Math.floor(rnd() * 40); i > 0; i--) cells[`${pick(rnd, COLS)}${pick(rnd, ROWS)}`] = randomCell(rnd, names, true)
      return { name, cells }
    }),
  }
}

/** A seeded step: cell edits (a value, a formula, a format alone, a clear) or a structure operation. */
type Step = { edits: CellEdit[] } | { op: StructureOp }

function randomSteps(seed: number, book: BookInput, count: number): Step[] {
  const rnd = mulberry32(seed)
  const names = book.sheets.map((s) => s.name)
  const steps: Step[] = []
  for (let i = 0; i < count; i++) {
    const s = pick(rnd, names)
    const x = rnd()
    if (x < 0.6) {
      steps.push({ edits: Array.from({ length: 1 + Math.floor(rnd() * 4) }, () => ({ sheet: s, cell: `${pick(rnd, COLS)}${pick(rnd, ROWS)}`, input: rnd() < 0.2 ? null : randomCell(rnd, names, !!book.styles?.length) })) })
      continue
    }
    const fresh = NAMES.find((n) => !names.includes(n)) ?? 'New'
    const op: StructureOp = pick(rnd, [
      { op: 'insertRows', sheet: s, row: pick(rnd, [1, 2, 10]), count: 1 + Math.floor(rnd() * 2) },
      { op: 'deleteRows', sheet: s, row: pick(rnd, [1, 2, 10]), count: 1 },
      { op: 'insertColumns', sheet: s, column: pick(rnd, ['A', 'B', 'Z']), count: 1 },
      { op: 'deleteColumns', sheet: s, column: pick(rnd, ['A', 'B', 'AA']), count: 1 },
      { op: 'renameSheet', sheet: s, name: fresh },
      { op: 'moveSheet', sheet: s },
      { op: 'insertSheet', name: fresh, before: s },
      { op: 'moveRange', from: { sheet: s, range: 'A1:B3' }, to: { sheet: s, cell: 'AA9' } },
    ] satisfies StructureOp[])
    if (op.op === 'renameSheet') names[names.indexOf(s)] = fresh
    if (op.op === 'insertSheet') names.push(fresh)
    steps.push({ op })
  }
  return steps
}

/** 1024 lines and more: lines hashed in several portions, a portion ending at a sheet's end and within a sheet. */
const portions: BookInput = { sheets: [sheet('P', Object.fromEntries(Array.from({ length: 2049 }, (_, i) => [`A${i + 1}`, { value: i % 7 ? i / 7 : `ы${i}` }]))), sheet('O', Object.fromEntries(Array.from({ length: 1024 }, (_, i) => [`B${i + 1}`, { value: i }])))] }

const BOOKS: [string, BookInput][] = [...EDGE, ['portions of lines', portions], ['the synthetic book', synthBook({ seed: 1, rows: 340 }).book], ...Array.from({ length: 24 }, (_, i): [string, BookInput] => [`seeded book ${i + 1}`, randomBook(i + 1)])]

// ---------- the tests ----------

const pools: { close(): Promise<void> }[] = []
afterAll(() => Promise.all(pools.map((p) => p.close())))

describe('hash and the canonical text', () => {
  it('a book keeps the hash recorded before streaming canonicalization', () => {
    // Recorded with commit 0b33695 (a Map of values, then one canonical text), before the hash was written as the cells are read.
    const expected: Record<string, string> = {
      'sheet and address order': '97affb292d7cc1d36a463be97956dbdb64951b046d4eaa090f23cac06da1f5de',
      values: '1de578aaad72b0a2904b65370abb4d351199edd6f785872f6cde267b4323a4ac',
      'spills over cells with only a format or a style': '41ff53c81ca86f86b225f911faa3d674ad77b916ad9d0a33d17b5becb117a608',
      'large spills and their fills': 'ac0511d0755ef98688f60baa26d274a83d71ce44dd0f40a2d002b2fa4b969066',
      'empty sheets and empty cells': '0b0116866572d38779b455e47acf9441f54f5486e3215d0e1dc4c9ca86adf17c',
      'values in ru-RU and the 1904 system': '7c4f769380ac8563651b6335d9d261f17955d43924a37c77785398fb4373e214',
      'large spills in ru-RU': '72302c93c6df0fb75e4b654b8d7f131043df8dd0c348820cca76c748c9b56915',
    }
    expect(EDGE.map(([name, book]) => [name, compute(book, { clock: CLOCK }).hash])).toEqual(Object.entries(expected))
  })

  it('en-US given or not: the same hash', () => {
    const book = EDGE[1][1]
    expect(compute(withSettings(book, { locale: 'en-US' }), { clock: CLOCK }).hash).toBe(compute(book, { clock: CLOCK }).hash)
    expect(compute(withSettings(book, { dateSystem: 1900 }), { clock: CLOCK }).hash).toBe(compute(book, { clock: CLOCK }).hash)
  })

  it('a cell with only a format or a style inside a spill is hashed once, as its spilled value', () => {
    const book = EDGE[2][1]
    // A2 and A3 are input cells that do not block the spill, so the result's cells list them twice: as input and as spilled
    const state = createState(book, { now: clockSerial(CLOCK), seed: CLOCK.seed })
    expect([1, 2].map((r) => state.spills.occupied.has(0, r, 0))).toEqual([false, false])
    expect([...state.spills.on.values()].find((x) => x.r1 === 0 && x.c1 === 0)).toMatchObject({ h: 3, w: 1 })
    const r = compute(book, { clock: CLOCK })
    expect([r.sheets[0].cells.A2, r.sheets[0].cells.A3]).toMatchObject([{ value: 2, spillOf: 'A1' }, { value: 3, spillOf: 'A1' }])
    expect(r.hash).toBe(sha256(referenceText(r, book, [])))
  })

  it('fills are hashed by sheet name, then top-left cell: a spill\'s fill on the right before its fill below, across spills', () => {
    const book = EDGE[3][1]
    const fills = fillsOf(book, 'closures')
    // the case covers both fills of one spill (G1) and fills of three spills on one sheet, in an order other than the hash's
    expect(fills.filter((f) => f.sheet === 'T').map((f) => f.range).sort()).toEqual(['C4:C1048576', 'E2:E100000', 'G2:KT300', 'H1:KT1'])
    const r = compute(book, { clock: CLOCK })
    expect(referenceText(r, book, fills).split('\n').filter((l) => /^"[^"]*"![A-Z]+[0-9]+:/.test(l)).map((l) => l.split(' ')[0])).toEqual(['"Fill"!B1:B1', '"Fill"!A2:B100000', '"Fill"!D2:D100000', '"T"!H1:KT1', '"T"!E2:E100000', '"T"!G2:KT300', '"T"!C4:C1048576', '"a"!A2:A100000'])
    expect(r.hash).toBe(sha256(referenceText(r, book, fills)))
  })
})

describe.each(['closures', 'eval'] as const)('execution %s', (execution: Execution) => {
  it('a synchronous session: books, then cell edits and structure operations', () => {
    BOOKS.forEach(([name, book], i) => {
      const s = openSession(book, { clock: CLOCK, execution })
      agree(s, execution, name)
      randomSteps(100 + i, book, 6).forEach((step, j) => {
        if ('edits' in step) s.setCells(step.edits)
        else s.restructure(step.op)
        agree(s, execution, `${name}, step ${j + 1}: ${JSON.stringify(step)}`)
      })
    })
  }, 300_000)

  it('in the same thread and in a thread, after the same steps', async () => {
    const calc = createInlineCalc({ sliceMs: 2 })
    const p = await pool({ threads: 2 })
    pools.push(p)
    for (const [i, [name, book]] of BOOKS.slice(0, EDGE.length + 8).entries()) {
      const s = openSession(book, { clock: CLOCK, execution })
      const [a, b] = [await calc.openSession(book, { clock: CLOCK, execution }), await p.openSession(book, { clock: CLOCK, execution })]
      const check = async (label: string) => {
        const h = s.hash()
        expect([await a.hash(), (await a.result()).hash, await b.hash(), (await b.result()).hash], label).toEqual([h, h, h, h])
      }
      await check(name)
      for (const [j, step] of randomSteps(100 + i, book, 4).entries()) {
        if ('edits' in step) {
          s.setCells(step.edits)
          await a.setCells(step.edits)
          await b.setCells(step.edits)
        } else {
          s.restructure(step.op)
          await a.restructure(step.op)
          await b.restructure(step.op)
        }
        await check(`${name}, step ${j + 1}`)
      }
      await Promise.all([a.close(), b.close()])
    }
  }, 300_000)
})
