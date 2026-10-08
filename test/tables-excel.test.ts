/**
 * Tables against Excel (`test/excel/tables/<set>.json`, from `run.ts tables <set>`; ru-RU sets in `ru-RU/`): each stored
 * case is its generator's current case, Excel loaded its book, and the engine refuses the steps Excel refused, ends with
 * the tables Excel lists, the values Excel shows in the rectangles and the same formulas, read as stored text. The
 * held-out sets (`-held`) keep the answers predicted before Excel ran them.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredTables } from '../bench/coverage/tables.js'
import type { TableSet } from '../bench/coverage/tables/cases.js'
import { differences } from './tables-cases.js'

const dir = new URL('./excel/tables/', import.meta.url)
const sets = (sub: string) => {
  const d = new URL(sub, dir)
  return existsSync(d) ? readdirSync(d).filter((f) => f.endsWith('.json')).map((f) => ({ name: f.slice(0, -5), sub, ...(JSON.parse(readFileSync(new URL(f, d), 'utf8')) as StoredTables) })) : []
}
const SETS = [...sets(''), ...sets('ru-RU/')]

/** `<set>:<case>` → why the engine's answer differs from Excel's (each has a line in the README's Known differences). */
const KNOWN: Readonly<Record<string, string>> = {}

describe.each(SETS)('tables $sub$name', (set) => {
  const mod = import(`../bench/coverage/tables/${set.name}.ts`) as Promise<TableSet>
  it.each(set.cases.map((c) => [c.id, c] as const))('%s', async (id, c) => {
    const now = (await mod).buildTableCases().find((x) => x.id === id)
    expect(now, 'the stored case is the generator\'s current one').toEqual(c)
    const diff = differences(c, set.answers[id])
    const known = KNOWN[`${set.name}:${id}`]
    if (known) expect(diff.length, `${id} is known to differ: ${known}`).toBeGreaterThan(0)
    else expect(diff).toEqual([])
  })
})
