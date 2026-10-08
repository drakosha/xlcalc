/**
 * Formulas of the old kind against Excel 365 (`test/excel/legacy/`): `fromLegacy` gives the formula Excel shows for
 * each (`formula2`, `@` where Excel puts it), and the engine computes it to the value Excel computed for the old
 * formula, at the row it stood in (implicit intersection takes the row). The sets: each whitelisted function with a
 * range and an expression in each argument (`params`), the same nested where arrays are computed (`nesting`, `blocks`), calls
 * whose answer depends on their arguments (`calls`), and the coverage measurement as Excel computed it with formulas of
 * the old kind (`measurement.json`).
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { Val } from '../bench/coverage/formulajs.js'
import type { Probe, SheetCells } from '../bench/coverage/probes.js'
import { FUNCTIONS } from '../src/functions.js'
import { compute, fromLegacy } from '../src/index.js'
import { canonical, parse } from '../src/syntax.js'
import { CASES, CLOCK, engineFormula, excelBook, matches, show, toVal, verdict } from './excel-cases.js'
import { caseBook } from './excel-sets.js'

interface Stored {
  cases: Probe[]
  values: Record<string, Val>
  formulas: Record<string, string>
  rows: Record<string, number>
}

const load = <T>(name: string) => JSON.parse(readFileSync(new URL(`./excel/legacy/${name}.json`, import.meta.url), 'utf8')) as T

/** Excel's text as canonical text; as it is where the engine has no syntax for it. */
function canon(f: string): string {
  try {
    return canonical(parse(engineFormula(f.replace(/^=/, ''))))
  } catch {
    return f
  }
}

/**
 * Cases whose value the converted formula does not give, and why (the dynamic text computed in Excel too: the spill set
 * `legacy-shown`). The old and the dynamic kind part in Excel itself: a range where an argument takes one value and
 * Excel shows no `@` is an error in the old formula and lifted in the dynamic one. And Excel computes a few dynamic
 * formulas otherwise than the engine: an error as NETWORKDAYS.INTL's weekend, XIRR's guess given a range (`#VALUE!`
 * in Excel), the functions of the former Analysis ToolPak and GROWTH, TREND over an array inside SUMPRODUCT.
 */
const PARTED: Readonly<Record<string, string>> = {
  'params:GROWTH#r4': 'Excel parts the two kinds',
  'params:MATCH#r3': 'Excel parts the two kinds',
  'params:TREND#r4': 'Excel parts the two kinds',
  'params:NETWORKDAYS.INTL#r3': 'an error as the weekend: #VALUE! in Excel, #NUM! here',
  'params:NETWORKDAYS.INTL#e3': 'an error as the weekend: #VALUE! in Excel, #NUM! here',
  'params:XIRR#r3': "XIRR's guess given a range: #VALUE! in Excel, lifted here",
  'nesting:MATCH#array-r3': 'Excel parts the two kinds',
  'nesting:ACCRINT#array-e1': 'inside SUMPRODUCT Excel does not lift the call',
  'nesting:ACCRINT#array-e2': 'inside SUMPRODUCT Excel does not lift the call',
  'nesting:CUMIPMT#array-e1': 'inside SUMPRODUCT Excel does not lift the call',
  'nesting:GROWTH#array-r4': 'inside SUMPRODUCT Excel does not lift the call',
  'nesting:PRICE#array-e1': 'inside SUMPRODUCT Excel does not lift the call',
  'nesting:TREND#array-r4': 'inside SUMPRODUCT Excel does not lift the call',
  'nesting:XIRR#array-r3': 'inside SUMPRODUCT Excel does not lift the call',
}

/** Volatile calls: their values are not compared. */
const VOLATILE = /\b(RAND|RANDBETWEEN|RANDARRAY|NOW|TODAY)\(/

/** The cases of a run in the books they stood in: a new book where the rows start again (a book split on a refused formula too). */
function books(cases: Probe[], rows: Record<string, number>): Probe[][] {
  const out: Probe[][] = []
  for (const c of cases) {
    if (!out.length || rows[c.id] <= rows[out.at(-1)!.at(-1)!.id]) out.push([])
    out.at(-1)!.push(c)
  }
  return out
}

/** A case book with each case at its own row (`caseBook` puts them one after another). */
function rowBook(cases: Probe[], rows: Record<string, number>, sheets?: SheetCells) {
  const book = caseBook([], sheets)
  const c = book.sheets.at(-1)!
  c.cells = Object.fromEntries(cases.map((x) => [`A${rows[x.id]}`, { formula: x.formula }]))
  return book
}

for (const name of ['params', 'nesting', 'calls', 'blocks']) {
  const set = load<Stored>(name)
  const sheets = ((await import(`../bench/coverage/legacy/${name}.ts`)) as { SHEETS?: SheetCells }).SHEETS

  describe(`formulas of the old kind against Excel 365: set ${name}`, () => {
    it("fromLegacy gives Excel's text, @ where Excel puts it", () => {
      const off = set.cases
        .map((c) => ({ c, excel: canon(set.formulas[c.id]), ours: fromLegacy(engineFormula(c.formula)) }))
        .filter((o) => o.excel !== o.ours)
        .map((o) => `${o.c.id} ${o.c.formula}: Excel ${o.excel}, fromLegacy ${o.ours}`)
      expect(off).toEqual([])
    })

    it("the converted formula computes Excel's value of the old one at its row", () => {
      const off: string[] = []
      for (const part of books(set.cases, set.rows)) {
        const probes = part.map((c) => ({ ...c, formula: fromLegacy(engineFormula(c.formula)) }))
        const cells = compute(rowBook(probes, set.rows, sheets), { clock: CLOCK }).sheets.at(-1)!.cells
        part.forEach((c) => {
          if (VOLATILE.test(c.formula)) return
          const engine = toVal(cells[`A${set.rows[c.id]}`].value)
          if (matches(c.fn, set.values[c.id], engine) === !!PARTED[`${name}:${c.id}`]) off.push(`${c.id} ${c.formula}: Excel ${show(set.values[c.id])}, engine ${show(engine)}`)
        })
      }
      expect(off).toEqual([])
    })
  })
}

describe('the coverage measurement computed as formulas of the old kind', () => {
  const legacy = load<{ values: Record<string, Val> }>('measurement').values
  const PICKED = new Set((JSON.parse(readFileSync(new URL('./excel/coverage/cases.json', import.meta.url), 'utf8')) as { cases: { id: string; pick?: unknown }[] }).cases.filter((c) => c.pick !== undefined).map((c) => c.id))
  // the parts Excel computed: whole functions, up to 400 cases each, in the order of the cases
  const parts: (typeof CASES)[] = []
  const cases = CASES.filter((c) => c.fn in FUNCTIONS)
  for (let i = 0; i < CASES.length; ) {
    let j = i
    while (j < CASES.length && CASES[j].fn === CASES[i].fn) j++
    if (!parts.length || parts.at(-1)!.length + (j - i) > 400) parts.push([])
    parts.at(-1)!.push(...CASES.slice(i, j))
    i = j
  }

  it('every case converted computes the value Excel gave the old formula at its row', () => {
    const off: string[] = []
    for (const part of parts) {
      // a case with an array answer was an array formula of the old kind on its cell, the others formulas of the old kind
      const book = excelBook(part)
      const cells = book.sheets.at(-1)!.cells
      part.forEach((c, i) => (cells[`A${i + 1}`] = PICKED.has(c.id) ? { formula: engineFormula(c.formula), array: `A${i + 1}` } : { formula: fromLegacy(engineFormula(c.formula)) }))
      const out = compute(book, { clock: CLOCK }).sheets.at(-1)!.cells
      part.forEach((c, i) => {
        if (!cases.includes(c)) return
        const engine = toVal(out[`A${i + 1}`].value)
        const v = verdict(`measurement:${c.id}`, c.fn, legacy[c.id], engine)
        if (!v.ok) off.push(`${c.id} ${c.formula}: Excel ${show(legacy[c.id])}, engine ${show(engine)}${v.why ? ` (${v.why})` : ''}`)
      })
    }
    expect(off).toEqual([])
  })
})

describe('fromLegacy', () => {
  it('puts @ before a name by what its formula gives, seen from the formula\'s sheet', () => {
    const names = [
      { name: 'r', ref: "'S'!$A$1:$A$9" },
      { name: 'k', ref: '0.2' },
      { name: 'a', ref: '{1,2}' },
      { name: 'r', ref: "'S'!$A$1", scope: 'T' },
    ]
    expect(fromLegacy('r*2+k+a+SUM(r)', { names })).toBe('@r*2+k+@a+SUM(r)')
    expect(fromLegacy('r*2', { names, sheet: 'T' })).toBe('r*2')
    expect(fromLegacy('r*2')).toBe('r*2')
  })

  it('returns a formula it does not read as it is', () => {
    expect(fromLegacy('SUM(A1:A3')).toBe('SUM(A1:A3')
  })

  it('takes the reference operators whole: the reference they answer gets @ where one value goes', () => {
    expect(fromLegacy('A1:A3 B2:C2')).toBe('@A1:A3 B2:C2')
    expect(fromLegacy('SUM(A1:A3 A2:C2)')).toBe('SUM(A1:A3 A2:C2)')
    expect(fromLegacy('SUM(A1:INDEX(B1:B9,3))')).toBe('SUM(A1:INDEX(B1:B9,3))')
  })
})
