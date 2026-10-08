/**
 * Tables against Excel for Windows (`test/excel/win/tables/`, `bench/windows`): the engine refuses the steps Excel for
 * Windows refused, ends with the tables it lists, the values it shows in the rectangles and the same formulas, as
 * `tables-excel.test.ts` checks the Mac's. Known differences are pinned in `WIN_TABLES` with the reason.
 */
import { describe, expect, it } from 'vitest'
import type { StoredTables } from '../bench/coverage/tables.js'
import type { TableSet } from '../bench/coverage/tables/cases.js'
import { winFiles } from './excel-win-cases.js'
import { differences } from './tables-cases.js'

const SETS = [...winFiles<StoredTables>('tables/').map((s) => ({ ...s, sub: '' })), ...winFiles<StoredTables>('tables/ru-RU/').map((s) => ({ ...s, sub: 'ru-RU/' }))]

/** `<set>:<case>` → why the engine's answer differs from Excel for Windows's. */
const WIN_TABLES: Readonly<Record<string, string>> = {}

describe.each(SETS)('tables $sub$name against Excel for Windows', (set) => {
  const mod = import(`../bench/coverage/tables/${set.name}.ts`) as Promise<TableSet>
  it.each(set.cases.map((c) => [c.id, c] as const))('%s', async (id, c) => {
    expect((await mod).buildTableCases().find((x) => x.id === id), "the stored case is the generator's current one").toEqual(c)
    const diff = differences(c, set.answers[id])
    const known = WIN_TABLES[`${set.name}:${id}`]
    if (known) expect(diff.length, `${id} is known to differ: ${known}`).toBeGreaterThan(0)
    else expect(diff).toEqual([])
  })
})

describe('known differences from Excel for Windows: tables', () => {
  it('every WIN_TABLES entry names a stored case', () => {
    const keys = new Set(SETS.flatMap((s) => s.cases.map((c) => `${s.name}:${c.id}`)))
    expect(Object.keys(WIN_TABLES).filter((k) => !keys.has(k))).toEqual([])
  })
})
