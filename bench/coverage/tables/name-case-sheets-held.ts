/**
 * Held out, after the set `name-case-held`: whether sheets and tables compare their names as defined names do, by the
 * letters JavaScript's upper case changes otherwise than one for one (`ſ` to `S`, `ß` to `SS`, `ﬁ` to `FI`, `ı` to `I`;
 * `İ` stays). A sheet is renamed through Excel's object model to such a name beside a sheet of the other spelling, which
 * Excel refuses if it is the same name; a formula over a table is typed in the other spelling of its name, which gives
 * #NAME? if it is another. Predicted by the engine with the rule of `name-case-held` (`caseKey`) before Excel ran it
 * (`name-case-sheets-held.predicted.json`).
 */
import { PAIRS } from './name-case-held.js'
import { book, on, table, type TableCase, type TableStep } from './cases.js'

const typed = (cell: string, formula: string): TableStep => ({ edit: { sheet: 'S', cell, value: { formula } } })

export function buildTableCases(): TableCase[] {
  const cases: TableCase[] = []
  for (const [id, word, upper] of PAIRS) {
    cases.push({
      id: `sheet-${id}`,
      about: `the sheet X renamed ${word} beside the sheet ${upper}: refused if it is the same name`,
      book: book(['Q', { A1: '=X!A1*2' }], [upper, { A1: 3 }], ['X', { A1: 5 }]),
      steps: [{ op: { op: 'renameSheet', sheet: 'X', name: word } }],
      read: on('Q', 'A1'),
    })
    const t = table(`T${word}`, 'A1', ['x'], [[2], [3]])
    cases.push({
      id: `table-${id}`,
      about: `a table T${word}, formulas over T${upper} and T${word} typed`,
      book: book(['S', {}, t]),
      steps: [typed('C1', `ROWS(T${upper})`), typed('C2', `ROWS(T${word})`)],
      read: on('S', 'A1:C3'),
    })
  }
  return cases
}
