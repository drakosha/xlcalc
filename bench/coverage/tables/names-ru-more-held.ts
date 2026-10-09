/**
 * Held out: Excel's rules for defined names in ru-RU (`run.ts tables names-ru-more-held ru`), predicted by the engine
 * before Excel ran them (`names-ru-more-held.predicted.json`). Names are added through Excel's object model
 * (`Names.Add`), which refuses a name Excel does not take, and formulas are typed in Russian after. The books with a
 * name in the file measure why Excel ru-RU does not load a file with the name ИСТИНА (`names-ru-held` bool-word): the
 * name alone, in another case, ЛОЖЬ, the word in a formula without the name, the word inside longer names.
 */
import type { BookInput, NameInput } from '../../../src/contract.js'
import { book, on, type CellSpec, type TableCase, type TableStep } from './cases.js'

export const local = true

const ru = (b: BookInput): BookInput => ({ ...b, locale: 'ru-RU' })
const typed = (cell: string, text: string, sheet = 'S'): TableStep => ({ edit: { sheet, cell, value: { local: text } } })
const added = (name: string, ref: string): TableStep => ({ name: { name, ref } })

/** Sheet S with A1 = 100, the cells given and the names, in ru-RU. */
const named = (names: NameInput[], cells: Record<string, CellSpec> = {}) => ru({ ...book(['S', { A1: 100, ...cells }]), names })

export function buildTableCases(): TableCase[] {
  const cases: TableCase[] = []
  const add = (id: string, about: string, b: BookInput, steps: TableStep[], read = on('S', 'A1:C3')) => cases.push({ id, about, book: b, ...(steps.length ? { steps } : {}), read })

  add('add-bool', 'the booleans of ru-RU added as names (ИСТИНА, ЛОЖЬ), and TRUE and FALSE, then ИСТИНА and ЛОЖЬ typed in Russian', named([]), [added('ИСТИНА', '11'), added('ЛОЖЬ', '12'), added('TRUE', '1'), added('FALSE', '2'), typed('C1', '=ИСТИНА+1'), typed('C2', '=ЛОЖЬ+1')])
  add('add-cyrillic', 'a Cyrillic name added, typed in Russian in another case', named([]), [added('Ставка', '0.2'), typed('C1', '=Ставка*10'), typed('C2', '=СУММ(ставка;1)')])
  add('add-shape', 'names shaped like a cell or R1C1: in Cyrillic letters names (АБ12, К2С3, С, К), in Latin ones refused (AB12, R1C1, R)', named([]), [added('АБ12', '1'), added('К2С3', '10'), added('С', '100'), added('К', '1000'), added('AB12', '2'), added('R1C1', '3'), added('R', '4'), typed('C1', '=АБ12+К2С3+С+К')])
  add('add-function', "function names as names in ru-RU: СУММ, Сумма and SUM, typed without (", named([]), [added('СУММ', '7'), added('Сумма', '5'), added('SUM', '3'), typed('C1', '=СУММ(СУММ;Сумма;SUM)')])
  add('add-chars', 'a question mark and a backslash inside a name (Став?ка, Став\\ка)', named([]), [added('Став?ка', '5'), added('Став\\ка', '6'), typed('C1', '=Став?ка*2'), typed('C2', '=Став\\ка*2')])
  add('file-bool', 'a file with the name ИСТИНА, which no formula uses', named([{ name: 'ИСТИНА', ref: '11' }]), [])
  add('file-bool-lower', 'a file with the name истина, which no formula uses', named([{ name: 'истина', ref: '11' }]), [])
  add('file-false', 'a file with the name ЛОЖЬ, which no formula uses', named([{ name: 'ЛОЖЬ', ref: '12' }]), [])
  add('file-bool-formula', 'a file with the word ИСТИНА in a formula and no such name: an unknown name', named([], { B1: '=ИСТИНА+1' }), [])
  add('file-bool-inside', 'a file with names that hold the word: ИСТИНА_1, НЕИСТИНА, ИСТИНАЛОЖЬ', named(['ИСТИНА_1', 'НЕИСТИНА', 'ИСТИНАЛОЖЬ'].map((name, i) => ({ name, ref: String(10 ** i) })), { B1: '=ИСТИНА_1+НЕИСТИНА+ИСТИНАЛОЖЬ' }), [])
  return cases
}
