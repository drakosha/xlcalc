/**
 * Excel-verified case sets beyond the coverage measurement: `bench/coverage/sets/<name>.ts` builds the
 * cases, `npx tsx bench/coverage/run.ts probes <name>` stores Excel's answers (en-US locale, English UI) in
 * `docs/excel-cases/<name>.json`. One file per set, so parallel work adds cases without touching shared files.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import type { Val } from '../bench/coverage/formulajs.js'
import type { Probe, SheetCells } from '../bench/coverage/probes.js'
import type { BookInput, CellInput } from '../src/contract.js'
import { dataSheet, engineFormula } from './excel-cases.js'

export interface StoredSet {
  name: string
  cases: Probe[]
  values: Record<string, Val>
}

const dir = new URL('../docs/excel-cases/', import.meta.url)

export const SETS: StoredSet[] = existsSync(dir)
  ? readdirSync(dir)
      .filter((f) => f.endsWith('.json'))
      .sort()
      .map((f) => ({ name: f.slice(0, -5), ...(JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as Omit<StoredSet, 'name'>) }))
  : []

/** Functions that some set has cases for (`fn` of its cases). */
export const SET_FUNCTIONS: ReadonlySet<string> = new Set(SETS.flatMap((s) => s.cases.flatMap((c) => (c.fn ? [c.fn] : []))))

/** A case set's generator. */
export const loadSet = async (name: string) => (await import(`../bench/coverage/sets/${name}.ts`)) as { buildProbes: () => Probe[]; SHEETS?: SheetCells }

const input = (v: SheetCells[string][string]): CellInput => (typeof v === 'object' ? { formula: engineFormula(v.formula) } : { value: v })

/** Cases on sheet C (the last one) in a book as sent to Excel: the measurement's data sheet D and the set's own sheets beside them. */
export function caseBook(probes: readonly Probe[], sheets: SheetCells = {}): BookInput {
  const own = Object.entries(sheets).map(([name, cells]) => ({ name, cells: Object.fromEntries(Object.entries(cells).map(([a, v]) => [a, input(v)])) }))
  const cases = Object.fromEntries(probes.map((p, i): [string, CellInput] => [`A${i + 1}`, { formula: engineFormula(p.formula) }]))
  return { sheets: [dataSheet(), ...own, { name: 'C', cells: cases }] }
}
