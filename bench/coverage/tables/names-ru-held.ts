/**
 * Held out: defined names in letters of any script in Excel ru-RU (`run.ts tables names-ru-held ru`), predicted by the
 * engine before Excel computed them (`names-ru-held.predicted.json`). Formulas typed in Russian: a name is written as
 * it is, a Russian function name followed by `(` is the function (`СУММ(Сумма;1)`), the same word without it is a name;
 * a name that is a Russian boolean (ИСТИНА) is the name in the file's formulas, while typed in Russian the word is the
 * boolean, as `toInvariant` reads it. Doubtful names have cases of their own: a name Excel would not take makes it
 * repair the file (the case is then `rejected`). Excel ru-RU did not open the book with the name ИСТИНА, which the
 * writer now leaves out of a ru-RU book's file, so that book does not go into a file whole any more and is not run again.
 */
import type { BookInput, NameInput } from '../../../src/contract.js'
import { book, on, sales, type CellSpec, type TableCase, type TableStep } from './cases.js'

export const local = true

const ru = (b: BookInput): BookInput => ({ ...b, locale: 'ru-RU' })
const typed = (cell: string, text: string, sheet = 'S'): TableStep => ({ edit: { sheet, cell, value: { local: text } } })

/** Sheet S with A1 = 100, the cells given and the names, in ru-RU. */
const named = (names: NameInput[], cells: Record<string, CellSpec>) => ru({ ...book(['S', { A1: 100, ...cells }]), names })

export function buildTableCases(): TableCase[] {
  const cases: TableCase[] = []
  const add = (id: string, about: string, b: BookInput, steps: TableStep[], read = on('S', 'A1:C3')) => cases.push({ id, about, book: b, steps, read })

  add(
    'typed',
    'Cyrillic names in formulas of the file, shown in Russian, and typed in Russian in any case',
    named([{ name: 'Ставка', ref: '0.2' }, { name: 'НДС_2025', ref: "'S'!$A$1" }], { B1: '=Ставка*2', B2: '=SUM(Ставка,НДС_2025)' }),
    [typed('C1', '=Ставка*10'), typed('C2', '=СУММ(Ставка;НДС_2025;0,5)'), typed('C3', '=СУММ(ставка;1)')],
  )
  add('function-word', 'a name Сумма beside the function СУММ', named([{ name: 'Сумма', ref: '5' }], { B1: '=SUM(Сумма)' }), [typed('C1', '=СУММ(Сумма;1)'), typed('C2', '=Сумма*2')])
  add('summ', "a name СУММ, SUM's Russian name: without ( it is the name", named([{ name: 'СУММ', ref: '7' }], { B1: '=СУММ*2' }), [typed('C1', '=СУММ(СУММ;1)'), typed('C2', '=СУММ+1')])
  add('bool-word', 'a name ИСТИНА: the name in the formula of the file, the boolean TRUE when typed in Russian', named([{ name: 'ИСТИНА', ref: '11' }], { B1: '=ИСТИНА+1' }), [typed('C1', '=ИСТИНА+1')])
  add(
    'scope',
    "a sheet's own Cyrillic name on a sheet named in Cyrillic, typed in Russian",
    ru({
      sheets: [
        { name: 'S', cells: { B1: { formula: "'Лист'!Ставка*2" } } },
        { name: 'Лист', cells: { A1: { formula: 'Ставка' } } },
      ],
      names: [{ name: 'Ставка', ref: '0.2' }, { name: 'Ставка', ref: '0.3', scope: 'Лист' }],
    }),
    [typed('C1', '=Лист!Ставка+1')],
    [...on('S', 'A1:C1'), ...on('Лист', 'A1')],
  )
  add('cell-like', 'Cyrillic letters and digits are a name in ru-RU too (АБ12)', named([{ name: 'АБ12', ref: '1' }], {}), [typed('C1', '=АБ12*2')])
  add('r1c1-like', 'R1C1 of ru-RU is in Latin letters: Cyrillic К2С3, С2К3, С and К are names', named(['К2С3', 'С2К3', 'С', 'К'].map((name, i) => ({ name, ref: String(10 ** i) })), {}), [typed('C1', '=К2С3+С2К3+С+К')])
  add(
    'rename-cyrillic',
    'a table renamed in Cyrillic, then a formula over it typed in Russian',
    ru(book(['S', { J1: '=SUM(Sales[Qty])' }, sales()])),
    [{ op: { op: 'renameTable', table: 'Sales', name: 'Продажи' } }, typed('J2', '=СУММ(Продажи[Price])')],
    on('S', 'A1:J5'),
  )
  return cases
}
