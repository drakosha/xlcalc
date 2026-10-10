/**
 * Held out: which sheet names Excel puts in quotes in a reference (`quoteSheetIfNeeded`). A book of sheets named like
 * cells (A1, R1C1 and past the grid), booleans, digits, spaces and punctuation, letters of other scripts; the sheet `F`
 * refers to each sheet's A1, written as the engine stores it (always quoted), and Excel's formula text is read back
 * (`Range.Formula`, and `FormulaLocal` in a run in ru-RU). `PREDICTED`, each formula as the rule predicted Excel to write
 * it, was committed before Excel ran the set.
 */
import { on, type TableCase } from './cases.js'

export const local = true

/** The sheet names, distinct in any case: each sheet holds its number in A1. */
export const NAMES = [
  'Sales', 'Q1 data', '2024', 'A1', 'R1C1', 'R', 'C', 'RC', 'TRUE', 'Лист1', 'Таблица 1', 'Sheet-1', '_x', 'a.b', 'XFD1048576', 'XFE1', "O'Brien",
  'false', 'R2', 'C3', 'rc12', 'R0', 'R1048577', 'C16385', 'A1048577', 'A0', 'ab12', 'A01', 'Sheet1', 'ИСТИНА', 'ЛОЖЬ', 'Straße', '€uro', '1Q',
  '.a', 'x!y', 'R2D2', 'C3PO', 'Sheet#1', '$A', '日本', 'A1B', 'К1С1', 'Data_2024.v1', 'N', 'SUM', 'T', 'RC1C', '_1', 'a,b', 'x(1)', 'ZZZ1', 'XFD', 'A',
]

/**
 * The text Excel was predicted to write for `=<sheet>!A1`: quoted when the name is not letters, digits, `_` and `.` after a
 * letter or `_`, when it reads as a cell of the grid (`A1`, `XFD1048576`, `ab12`, `A01`; not `XFE1`, `A0`) or as an R1C1
 * reference by its letters (`R`, `C`, `RC`, `R1C1`, `R0`), or as a boolean (`TRUE`, `false`).
 */
export const PREDICTED: Readonly<Record<string, string>> = {
  Sales: '=Sales!A1',
  'Q1 data': "='Q1 data'!A1",
  2024: "='2024'!A1",
  A1: "='A1'!A1",
  R1C1: "='R1C1'!A1",
  R: "='R'!A1",
  C: "='C'!A1",
  RC: "='RC'!A1",
  TRUE: "='TRUE'!A1",
  Лист1: '=Лист1!A1',
  'Таблица 1': "='Таблица 1'!A1",
  'Sheet-1': "='Sheet-1'!A1",
  _x: '=_x!A1',
  'a.b': '=a.b!A1',
  XFD1048576: "='XFD1048576'!A1",
  XFE1: '=XFE1!A1',
  "O'Brien": "='O''Brien'!A1",
  false: "='false'!A1",
  R2: "='R2'!A1",
  C3: "='C3'!A1",
  rc12: "='rc12'!A1",
  R0: "='R0'!A1",
  R1048577: "='R1048577'!A1",
  C16385: "='C16385'!A1",
  A1048577: '=A1048577!A1',
  A0: '=A0!A1',
  ab12: "='ab12'!A1",
  A01: "='A01'!A1",
  Sheet1: '=Sheet1!A1',
  ИСТИНА: '=ИСТИНА!A1',
  ЛОЖЬ: '=ЛОЖЬ!A1',
  Straße: '=Straße!A1',
  '€uro': "='€uro'!A1",
  '1Q': "='1Q'!A1",
  '.a': "='.a'!A1",
  'x!y': "='x!y'!A1",
  R2D2: '=R2D2!A1',
  C3PO: '=C3PO!A1',
  'Sheet#1': "='Sheet#1'!A1",
  $A: "='$A'!A1",
  日本: '=日本!A1',
  A1B: '=A1B!A1',
  К1С1: '=К1С1!A1',
  'Data_2024.v1': '=Data_2024.v1!A1',
  N: '=N!A1',
  SUM: '=SUM!A1',
  T: '=T!A1',
  RC1C: '=RC1C!A1',
  _1: '=_1!A1',
  'a,b': "='a,b'!A1",
  'x(1)': "='x(1)'!A1",
  ZZZ1: '=ZZZ1!A1',
  XFD: '=XFD!A1',
  A: '=A!A1',
}

/** The cell of `F` that refers to the sheet `name` of `names`: one per name, down column A. */
export const cellOf = (name: string, names = NAMES) => `A${names.indexOf(name) + 1}`

/** The case of a list of sheet names: `F` refers to each sheet's A1, which holds the sheet's number. */
export function quoteCases(names: string[]): TableCase[] {
  const quoted = (name: string) => `'${name.replaceAll("'", "''")}'`
  const refs = Object.fromEntries(names.map((name) => [cellOf(name, names), { formula: `${quoted(name)}!A1` }]))
  const sheets = [{ name: 'F', cells: refs }, ...names.map((name, i) => ({ name, cells: { A1: { value: i + 1 } } }))]
  return [{ id: 'names', about: 'a reference to a sheet of each name, read back as Excel writes it', book: { sheets }, read: on('F', `A1:A${names.length}`) }]
}

export const buildTableCases = (): TableCase[] => quoteCases(NAMES)
