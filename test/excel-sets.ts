/**
 * Excel-verified case sets beyond the coverage measurement: `bench/coverage/sets/<name>.ts` builds the
 * cases, `npx tsx bench/coverage/run.ts probes <name>` stores Excel's answers (en-US locale, English UI) in
 * `docs/excel-cases/<name>.json`. One file per set, so parallel work adds cases without touching shared files.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import type { Val } from '../bench/coverage/formulajs.js'
import type { Probe } from '../bench/coverage/probes.js'

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
