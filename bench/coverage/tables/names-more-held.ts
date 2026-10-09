/**
 * Held out: Excel's rules for defined names in en-US, measured where they refuse, predicted by the engine before Excel
 * ran them (`names-more-held.predicted.json`). Names are added through Excel's object model (`Names.Add`), which
 * refuses a name Excel does not take, as the engine's `setNames` leaves it out; formulas typed after show whether a
 * formula reads the name, and the names Excel lists at the end are compared. A file holds the names the engine's
 * writer takes, `?` and `\` inside among them, so two books measure whether Excel loads such a file.
 */
import type { NameInput } from '../../../src/contract.js'
import { book, on, sales, type TableCase, type TableStep } from './cases.js'

const typed = (cell: string, formula: string, sheet = 'S'): TableStep => ({ edit: { sheet, cell, value: { formula } } })
const added = (name: string, ref: string, scope?: string): TableStep => ({ name: { name, ref, ...(scope ? { scope } : {}) } })
const S = () => book(['S', { A1: 100 }])

export function buildTableCases(): TableCase[] {
  const cases: TableCase[] = []
  const add = (id: string, about: string, b: TableCase['book'], steps: TableStep[], read = on('S', 'A1:C3')) => cases.push({ id, about, book: b, ...(steps.length ? { steps } : {}), read })
  const long = 'Я'.repeat(255)

  add('add-cyrillic', 'a Cyrillic name added, used in its case and in another', S(), [added('Ставка', '0.2'), typed('C1', 'Ставка*10'), typed('C2', 'ставка+1')])
  add(
    'add-scope',
    "a global name and a sheet's own of the same name added, the sheet's own on a sheet named in Cyrillic",
    { sheets: [{ name: 'S', cells: { A1: { value: 100 } } }, { name: 'Лист', cells: {} }] },
    [added('Ставка', '0.2'), added('Ставка', '0.3', 'Лист'), typed('C1', 'Ставка'), typed('C2', 'Лист!Ставка+1'), typed('A1', 'ставка', 'Лист')],
    [...on('S', 'A1:C3'), ...on('Лист', 'A1')],
  )
  add(
    'add-shape',
    'names shaped like a cell or R1C1: in Cyrillic letters names (АБ12, К2С3, С), in Latin ones refused (AB12, R2C3, R, c, RC, R1C); RR is no reference',
    S(),
    [added('АБ12', '1'), added('AB12', '2'), added('К2С3', '10'), added('R2C3', '20'), added('С', '100'), added('R', '200'), added('c', '300'), added('RC', '400'), added('R1C', '500'), added('RR', '1000'), typed('C1', 'АБ12'), typed('C2', 'К2С3+С*10'), typed('C3', 'RR')],
  )
  add('add-question', 'a question mark inside a name (Став?ка, Rate?) and first (?Ставка)', S(), [added('Став?ка', '5'), added('?Ставка', '6'), added('Rate?', '7'), typed('C1', 'Став?ка*2'), typed('C2', 'Rate?*2')])
  add('add-backslash', 'a backslash inside a name (Став\\ка) and first (\\Ставка)', S(), [added('Став\\ка', '5'), added('\\Ставка', '6'), typed('C1', 'Став\\ка*2'), typed('C2', '\\Ставка*2')])
  add(
    'add-chars',
    'a dot, an underscore and a digit inside or first (Став.ка, _Ставка, Ставка1; 1Ставка refused), a space or a hyphen inside refused',
    S(),
    [added('Став.ка', '1'), added('_Ставка', '2'), added('Ставка1', '3'), added('1Ставка', '4'), added('Став ка', '5'), added('Став-ка', '6'), typed('C1', 'Став.ка+_Ставка*10+Ставка1*100')],
  )
  add('add-bool', 'TRUE and FALSE refused in any case, the booleans of ru-RU are names in en-US', S(), [added('TRUE', '1'), added('false', '2'), added('ИСТИНА', '11'), added('ЛОЖЬ', '12'), typed('C1', 'ИСТИНА+ЛОЖЬ')])
  add('add-function', "function names as names: SUM, and СУММ and Сумма, SUM's name and a word in ru-RU", S(), [added('SUM', '3'), added('СУММ', '7'), added('Сумма', '5'), typed('C1', 'SUM(SUM,СУММ,Сумма)'), typed('C2', 'SUM*2')])
  add('add-long', 'a name of 255 letters, and one of 256 refused', S(), [added(long, '1'), added(`${long}Я`, '2'), typed('C1', `${long}+1`)])
  const table = book(['S', { J1: '=SUM(Sales[Qty])' }, sales()])
  add('add-table', "a table's name in another case is no defined name (SALES beside the table Sales); Продажи is", table, [added('SALES', '1'), added('Продажи', '2'), typed('J2', 'Продажи*2')], on('S', 'J1:J2'))
  const file = (name: string): TableCase['book'] => ({ ...S(), names: [{ name, ref: '5' } satisfies NameInput] })
  add('file-question', 'a file holding a name with a question mark inside (Став?ка): Excel loads it or repairs it', file('Став?ка'), [typed('C1', 'Став?ка*2')])
  add('file-backslash', 'a file holding a name with a backslash inside (Став\\ка): Excel loads it or repairs it', file('Став\\ка'), [typed('C1', 'Став\\ка*2')])
  return cases
}
