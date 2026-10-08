/**
 * Tables against Excel (`test/excel/tables/<set>.json`, from `run.ts tables <set>`; ru-RU sets in `ru-RU/`): each stored
 * case is its generator's current case, Excel loaded its book, and the engine refuses the steps Excel refused, ends with
 * the tables Excel lists, the values Excel shows in the rectangles and the same formulas, read as stored text. The
 * held-out sets (`-held`) keep the answers predicted before Excel ran them.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { compare } from '../bench/coverage/compare.js'
import type { StoredTables } from '../bench/coverage/tables.js'
import type { TableAnswer, TableSet } from '../bench/coverage/tables/cases.js'
import { engineAnswer, storedText } from './tables-cases.js'

const dir = new URL('./excel/tables/', import.meta.url)
const sets = (sub: string) => {
  const d = new URL(sub, dir)
  return existsSync(d) ? readdirSync(d).filter((f) => f.endsWith('.json')).map((f) => ({ name: f.slice(0, -5), sub, ...(JSON.parse(readFileSync(new URL(f, d), 'utf8')) as StoredTables) })) : []
}
const SETS = [...sets(''), ...sets('ru-RU/')]

/** `<set>:<case>` → why the engine's answer differs from Excel's (each has a line in the README's Known differences). */
const KNOWN: Readonly<Record<string, string>> = {}

/** What differs between the engine's answer and Excel's, as lines (empty when they agree). */
export function differences(c: Parameters<typeof engineAnswer>[0], x: TableAnswer): string[] {
  const out: string[] = []
  if (x.rejected) return ['Excel did not load the book']
  const got = engineAnswer(c)
  const refused = (r: Record<number, string> | undefined) => Object.keys(r ?? {}).join(',')
  if (refused(got.refused) !== refused(x.refused)) out.push(`refused steps: Excel ${refused(x.refused) || 'none'} (${Object.values(x.refused ?? {}).join('; ')}), engine ${refused(got.refused) || 'none'} (${Object.values(got.refused).join('; ')})`)
  const shape = (t: TableAnswer['tables'][number]) => JSON.stringify(t)
  const ex = x.tables.map(shape).sort()
  const en = got.tables.map(shape).sort()
  if (JSON.stringify(ex) !== JSON.stringify(en)) out.push(`tables: Excel ${ex.join(' ')}, engine ${en.join(' ')}`)
  for (const k of new Set([...Object.keys(x.values), ...Object.keys(got.values)])) {
    const want = x.values[k]
    const have = got.values[k]
    if (!want || !have) {
      if (want !== have) out.push(`${k}: Excel ${JSON.stringify(want ?? 'blank')}, engine ${JSON.stringify(have ?? 'blank')}`)
    } else if (!compare(want, have).ok) out.push(`${k}: Excel ${JSON.stringify(want)}, engine ${JSON.stringify(have)}`)
  }
  for (const [k, shown] of Object.entries(x.formulas)) {
    const [sheet, addr] = k.split('!')
    const stored = storedText(got.book, sheet, addr, shown)
    if (stored !== got.formulas[k]) out.push(`${k}: Excel ${shown} (${stored}), engine ${got.formulas[k] ?? 'no formula'}`)
  }
  return out
}

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
