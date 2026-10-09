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
const WIN_TABLES: Readonly<Record<string, string>> = {
  'ops-held:rename-cell': "a table renamed to a name Excel does not take (AB12): Excel's object model makes it one (AB12_), as on the Mac; renameTable refuses with table_name",
  'ops-held:rename-name': "a table renamed to a defined name (Rate): Excel's object model makes the name free (Rate_1), as on the Mac; renameTable refuses with table_name",
  'more:type-below': "a value typed right below a table: Excel's AutoCorrect extends the table by a row, as on the Mac; setCells leaves the table as it is",
  'more:type-calc': "a formula typed into an empty column of a table: Excel's AutoCorrect fills the column with it, as on the Mac; setCells puts it in its cell",
  'names-held:rename-digit': "a table renamed to a name Excel does not take (1Ставка): Excel's object model makes it one (_1Ставка), as on the Mac; renameTable refuses with table_name",
  'names-held:rename-name-case': "a table renamed to a defined name in another case (СТАВКА beside Ставка): Excel's object model makes the name free (СТАВКА_1), as on the Mac; renameTable refuses with table_name",
  'names-tables-held:rename-question-first': "a table renamed to a name Excel does not take (?Продажи): Excel's object model makes it one (_?Продажи), as on the Mac; renameTable refuses with table_name",
  'names-ru-more-held:add-bool': 'Excel ru-RU takes TRUE and FALSE as defined names, as on the Mac; the engine keeps formulas in en-US text, where they are the booleans, and refuses them in every locale',
  'names-ru-tables-held:rename-bool': "a table renamed ИСТИНА in ru-RU, a name Excel for Windows does not take: its object model makes one (ИСТИНА_); renameTable refuses with table_name",
  'names-ru-tables-held:rename-bool-lower': "a table renamed ложь in ru-RU, a name Excel for Windows does not take: its object model makes one (ложь_); renameTable refuses with table_name",
  'names-ru-tables-held:rename-true': 'a table renamed TRUE in ru-RU: Excel for Windows keeps the name; the engine keeps formulas in en-US text, where TRUE is the boolean, and refuses it in every locale',
}

describe.each(SETS)('tables $name ($sub) against Excel for Windows', (set) => {
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
