/**
 * Held out: defined names in letters of any script in Excel en-US, predicted by the engine before Excel computed them
 * (`names-held.predicted.json`, from `predict.ts`). The rules predicted are Excel's rules for names, which tables have
 * too: a letter of any script, `_` or `\` first, then letters, digits, `_` and `.`; a cell address and R1C1 in Latin
 * letters only (`АБ12`, `К2С3` and a Cyrillic `С` are names); names compared in any case. A name Excel would not take
 * makes it repair the file (the case is then `rejected`), so each doubtful name has a case of its own. A file cannot
 * hold a name the engine refuses, so refusals are measured through tables renamed, whose names follow the same rules.
 */
import type { NameInput, StructureOp } from '../../../src/contract.js'
import { book, on, sales, type CellSpec, type TableCase, type TableStep } from './cases.js'

const op = (o: StructureOp): TableStep => ({ op: o })
const typed = (cell: string, formula: string, sheet = 'S'): TableStep => ({ edit: { sheet, cell, value: { formula } } })

/** Sheet S with A1 = 100, the cells given and the names. */
const named = (names: NameInput[], cells: Record<string, CellSpec>) => ({ ...book(['S', { A1: 100, ...cells }]), names })

const OVER: Record<string, CellSpec> = { J1: '=SUM(Sales[Qty])', J2: '=ROWS(Sales)' }

export function buildTableCases(): TableCase[] {
  const cases: TableCase[] = []
  const add = (id: string, about: string, b: TableCase['book'], steps: TableStep[] = [], read = on('S', 'A1:C3')) => cases.push({ id, about, book: b, ...(steps.length ? { steps } : {}), read })
  const rename = (id: string, about: string, name: string, b = book(['S', OVER, sales()])) => add(id, about, b, [op({ op: 'renameTable', table: 'Sales', name })], on('S', 'A1:K5'))

  add(
    'cyrillic',
    'names in Cyrillic: a constant and a reference, in formulas of the file and typed',
    named([{ name: 'Ставка', ref: '0.2' }, { name: 'НДС_2025', ref: "'S'!$A$1" }], { B1: '=Ставка*2', B2: '=SUM(Ставка,НДС_2025)', B3: '=НДС_2025*Ставка' }),
    [typed('C1', 'Ставка*10'), typed('C2', 'SUM(НДС_2025,Ставка)')],
  )
  add(
    'any-case',
    'a Cyrillic name used in other cases is the same name (Excel may show it in the case it was defined in)',
    named([{ name: 'Ставка', ref: '0.2' }], { B1: '=ставка*2', B2: '=СТАВКА+1' }),
    [typed('C1', 'сТаВкА*10')],
  )
  add('cell-like', 'Cyrillic letters and digits are a name, not a cell (АБ12; aБ12 with a Latin a)', named([{ name: 'АБ12', ref: '1' }, { name: 'aБ12', ref: '2' }], { B1: '=АБ12+aБ12*10' }))
  add('r1c1-like', 'Cyrillic letters like R1C1 and like R and C alone are names (К2С3, Р1С1, С, К)', named(['К2С3', 'Р1С1', 'С', 'К'].map((name, i) => ({ name, ref: String(10 ** i) })), { B1: '=К2С3+Р1С1+С+К' }))
  add('letters', 'Ё, Latin and Cyrillic mixed (aБ), Greek (Τιμή)', named([{ name: 'Ё', ref: '1' }, { name: 'aБ', ref: '2' }, { name: 'Τιμή', ref: '3' }], { B1: '=Ё+aБ*10+Τιμή*100' }))
  add('yo', 'Ё and Е are two letters: Ёлка and Елка are two names', named([{ name: 'Ёлка', ref: '1' }, { name: 'Елка', ref: '2' }], { B1: '=Ёлка', B2: '=Елка' }))
  add('dot-underscore', 'a dot and an underscore inside, an underscore first', named([{ name: 'НДС.2025', ref: '1' }, { name: '_Ставка', ref: '2' }, { name: 'Ставка_1.2', ref: '3' }], { B1: '=НДС.2025+_Ставка*10+Ставка_1.2*100' }))
  add('backslash', 'a backslash first', named([{ name: '\\Ставка', ref: '5' }], { B1: '=\\Ставка*2' }))
  add(
    'scope',
    "a sheet's own Cyrillic name on a sheet named in Cyrillic, over a global one of the same name",
    {
      sheets: [
        { name: 'S', cells: { B1: { formula: "'Лист'!Ставка*2" }, B2: { formula: 'Ставка' } } },
        { name: 'Лист', cells: { A1: { formula: 'ставка' } } },
      ],
      names: [{ name: 'Ставка', ref: '0.2' }, { name: 'Ставка', ref: '0.3', scope: 'Лист' }],
    },
    [typed('C1', 'Лист!Ставка+1')],
    [...on('S', 'A1:C2'), ...on('Лист', 'A1')],
  )
  add('word-ru', "Russian function names are plain words in en-US: Сумма, and СУММ, SUM's name in ru-RU", named([{ name: 'Сумма', ref: '5' }, { name: 'СУММ', ref: '7' }], { B1: '=Сумма*2', B2: '=SUM(Сумма,СУММ)', B3: '=СУММ' }))
  add('word-en', "a name that is an English function's (SUM): the engine takes it, Excel may not", named([{ name: 'SUM', ref: '3' }], { B1: '=SUM*2', B2: '=SUM(SUM,1)' }))
  add('word-bool-ru', "the booleans of ru-RU are plain words in en-US (ИСТИНА, ЛОЖЬ)", named([{ name: 'ИСТИНА', ref: '11' }, { name: 'ЛОЖЬ', ref: '12' }], { B1: '=ИСТИНА+ЛОЖЬ' }))
  rename('rename-cyrillic', 'a table renamed in Cyrillic: the references follow', 'Продажи')
  rename('rename-cell-like', 'a table renamed to Cyrillic letters and digits (АБ12): not a cell, Excel takes it', 'АБ12')
  rename('rename-r1c1-like', 'a table renamed to Cyrillic letters like R1C1 (К2С3): Excel takes it', 'К2С3')
  rename('rename-digit', 'a table renamed to a name with a digit first (1Ставка): Excel refuses (its object model may make a name of it)', '1Ставка')
  const taken = book(['S', { ...OVER, J3: '=SUM(Ставка)' }, sales()])
  taken.names = [{ name: 'Ставка', ref: "'S'!$C$2" }]
  rename('rename-name-case', 'a table renamed to a defined name in another case (СТАВКА over Ставка): one namespace, Excel refuses (its object model may make the name free)', 'СТАВКА', taken)
  return cases
}
