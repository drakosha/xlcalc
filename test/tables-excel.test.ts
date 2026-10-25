/**
 * Tables against Excel (`test/excel/tables/<set>.json`, from `run.ts tables <set> [locale]`; a locale's other than
 * en-US in a directory of its tag, `ru-RU/`): each stored case is its generator's current case, Excel loaded its book,
 * and the engine refuses the steps Excel refused, ends with the tables Excel lists, the values Excel shows in the
 * rectangles and the same formulas, read as stored text. The held-out sets (`-held`) keep the answers predicted before
 * Excel ran them.
 */
import { describe, expect, it } from 'vitest'
import type { StoredTables } from '../bench/coverage/tables.js'
import type { TableSet } from '../bench/coverage/tables/cases.js'
import { localeDir } from '../bench/coverage/locales.js'
import { differences } from './tables-cases.js'
import { localeSets } from './stored-sets.js'
import './locales.js'

const SETS = localeSets<StoredTables>('tables').map((r) => ({ name: r.set, sub: localeDir(r.tag), tag: r.tag, ...r.stored }))

/** `<set>:<case>` → why the engine's answer differs from Excel's (each has a line in docs/limitations.md's Known differences). */
const KNOWN: Readonly<Record<string, string>> = {
  'ops-held:rename-cell': "a table renamed to a name Excel does not take (AB12): Excel's object model makes it one (AB12_), as its name box refuses it; renameTable refuses with table_name",
  'ops-held:rename-name': "a table renamed to a defined name (Rate): Excel's object model makes the name free (Rate_1); renameTable refuses with table_name",
  'more:type-below': "a value typed right below a table: Excel's AutoCorrect extends the table by a row; setCells leaves the table as it is",
  'more:type-calc': "a formula typed into an empty column of a table: Excel's AutoCorrect fills the column with it; setCells puts it in its cell",
  'paste-names-held:paste-not-multiple': "formulas pasted into part of a table's calculated column: Excel fills the whole column with them, its calculated column following the paste; copyRange pastes the cells alone",
  'paste-names-held:paste-multiple': "formulas pasted into part of a table's calculated column: Excel fills the whole column with them, its calculated column following the paste; copyRange pastes the cells alone",
  'names-held:rename-digit': "a table renamed to a name Excel does not take (1Ставка): Excel's object model makes it one (_1Ставка); renameTable refuses with table_name",
  'names-held:rename-name-case': "a table renamed to a defined name in another case (СТАВКА beside Ставка): Excel's object model makes the name free (СТАВКА_1); renameTable refuses with table_name",
  'names-tables-held:rename-question-first': "a table renamed to a name Excel does not take (?Продажи): Excel's object model makes it one (_?Продажи); renameTable refuses with table_name",
  'names-more-held:add-long': 'a formula over a name of 255 letters, typed through AppleScript, which cuts what it sets to 255 characters (the formula loses its end and gives #NAME?); Excel for Windows computes it as the engine does',
  'names-ru-more-held:add-bool': 'Excel ru-RU takes TRUE and FALSE as defined names; the engine keeps formulas in en-US text, where they are the booleans, and refuses them in every locale',
  'names-ru-tables-held:rename-bool': 'a table renamed ИСТИНА in ru-RU: Excel for Mac keeps the name, Excel for Windows makes ИСТИНА_ of it as of a defined name; renameTable refuses with table_name, as Windows',
  'names-ru-tables-held:rename-bool-lower': 'a table renamed ложь in ru-RU: Excel for Mac keeps the name, Excel for Windows makes ложь_ of it as of a defined name; renameTable refuses with table_name, as Windows',
  'names-ru-tables-held:rename-true': "a table renamed TRUE in ru-RU, a name Excel for Mac does not take: its object model makes one (TRUE_); renameTable refuses with table_name",
  'name-case-held:long-s-word': "the long s (ſ): Excel for Mac takes it for S in a defined name, a table's name and a sheet's (Names.Add of ſ beside S replaces it, a sheet is not renamed ſ beside S); the engine, as Excel for Windows, keeps ſ a letter of its own",
  'name-case-ru-held:long-s-word': "the long s (ſ): Excel for Mac takes it for S in a defined name, a table's name and a sheet's (Names.Add of ſ beside S replaces it, a sheet is not renamed ſ beside S); the engine, as Excel for Windows, keeps ſ a letter of its own",
  'name-case-ru-held:long-s-upper': "the long s (ſ): Excel for Mac takes it for S in a defined name, a table's name and a sheet's (Names.Add of ſ beside S replaces it, a sheet is not renamed ſ beside S); the engine, as Excel for Windows, keeps ſ a letter of its own",
  'name-case-held:long-s-upper': "the long s (ſ): Excel for Mac takes it for S in a defined name, a table's name and a sheet's (Names.Add of ſ beside S replaces it, a sheet is not renamed ſ beside S); the engine, as Excel for Windows, keeps ſ a letter of its own",
  'name-case-held:long-s-both': "the long s (ſ): Excel for Mac takes it for S in a defined name, a table's name and a sheet's (Names.Add of ſ beside S replaces it, a sheet is not renamed ſ beside S); the engine, as Excel for Windows, keeps ſ a letter of its own",
  'name-case-sheets-held:sheet-long-s': "the long s (ſ): Excel for Mac takes it for S in a defined name, a table's name and a sheet's (Names.Add of ſ beside S replaces it, a sheet is not renamed ſ beside S); the engine, as Excel for Windows, keeps ſ a letter of its own",
  'name-case-sheets-held:table-long-s': "the long s (ſ): Excel for Mac takes it for S in a defined name, a table's name and a sheet's (Names.Add of ſ beside S replaces it, a sheet is not renamed ſ beside S); the engine, as Excel for Windows, keeps ſ a letter of its own",
  'name-case-sheets-held:sheet-sharp-s': "a sheet renamed ß beside a sheet SS (ﬁ beside FI): Excel, for Mac and for Windows, compares sheet names letter by letter and takes it; the engine compares sheets as defined names and tables, where ß is SS, and refuses the name",
  'name-case-sheets-held:sheet-ligature-fi': "a sheet renamed ß beside a sheet SS (ﬁ beside FI): Excel, for Mac and for Windows, compares sheet names letter by letter and takes it; the engine compares sheets as defined names and tables, where ß is SS, and refuses the name",
  'name-values-held:qualified-delete-S': "a sheet deleted under references to its own name (S!loc): Excel writes the book's file name for the sheet ('book.xlsx'!loc), the engine leaves the text; the values are #REF! in both",
  'headers-ru-held:text-iso': 'text typed into a header cell that Excel reads as a date (2025-01-31): Excel names the column by the date as it shows it (31.01.2025); setCells takes text as text, and the column is named 2025-01-31',
}

describe.each(SETS)('tables $sub$name', (set) => {
  const mod = import(`../bench/coverage/tables/${set.name}.ts`) as Promise<TableSet>
  it.each(set.cases.map((c) => [c.id, c] as const))('%s', async (id, c) => {
    const now = (await mod).buildTableCases().find((x) => x.id === id)
    expect(now, 'the stored case is the generator\'s current one').toEqual(c)
    const diff = differences(c, set.answers[id], set.tag)
    const known = KNOWN[`${set.name}:${id}`]
    if (known) expect(diff.length, `${id} is known to differ: ${known}`).toBeGreaterThan(0)
    else expect(diff).toEqual([])
  })
})
