/**
 * Excel-verified case sets beyond the coverage measurement: `bench/coverage/sets/<name>.ts` builds the
 * cases, `npx tsx bench/coverage/run.ts probes <name>` stores Excel's answers (en-US locale, English UI) in
 * `test/excel/cases/<name>.json`. One file per set, so parallel work adds cases without touching shared files.
 */
import type { Val } from '../bench/coverage/formulajs.js'
import type { Probe, SheetCells } from '../bench/coverage/probes.js'
import type { BookInput, CellInput } from '../src/contract.js'
import { caseCells, dataSheet, engineFormula } from './excel-cases.js'
import { storedSets } from './stored-sets.js'

export interface StoredSet {
  name: string
  cases: Probe[]
  values: Record<string, Val>
}

export const SETS: StoredSet[] = storedSets<Omit<StoredSet, 'name'>>('cases')

/** The sets of match mode 3's syntax, which each level of `xlcalc/regex` answers differently: `regex-levels.test.ts` checks them. */
export const LEVELED_SETS: ReadonlySet<string> = new Set(['xmatch-regex-own-held', 'xmatch-regex-syntax', 'xmatch-regex-syntax-2'])

/** Functions that some set has cases for (`fn` of its cases). */
export const SET_FUNCTIONS: ReadonlySet<string> = new Set(SETS.flatMap((s) => s.cases.flatMap((c) => (c.fn ? [c.fn] : []))))

/** A case set's generator. */
export const loadSet = async (name: string) => (await import(`../bench/coverage/sets/${name}.ts`)) as { buildProbes: () => Probe[]; SHEETS?: SheetCells }

const input = (v: SheetCells[string][string]): CellInput => (typeof v === 'object' ? { formula: engineFormula(v.formula) } : { value: v })

/**
 * Cases on sheet C (the last one) in a book as sent to Excel: the measurement's data sheet D and the set's own
 * sheets beside them. As in Excel's book, every case has a neighbour on its right and the last one a value
 * below it, so an array answer cannot spill (`#SPILL!`).
 */
export function caseBook(probes: readonly Probe[], sheets: SheetCells = {}): BookInput {
  const own = Object.entries(sheets).map(([name, cells]) => ({ name, cells: Object.fromEntries(Object.entries(cells).map(([a, v]) => [a, input(v)])) }))
  return { sheets: [dataSheet(), ...own, { name: 'C', cells: caseCells(probes.map((p) => engineFormula(p.formula))) }] }
}
