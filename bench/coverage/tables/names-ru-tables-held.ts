/**
 * Held out: tables named by the rules of names in ru-RU (`run.ts tables names-ru-tables-held ru`), where Excel ru-RU
 * refuses a defined name ИСТИНА or ЛОЖЬ and takes TRUE (`names-ru-more-held`); predicted by the engine before Excel ran
 * them (`names-ru-tables-held.predicted.json`). Excel's object model makes a name of one it does not take.
 */
import { book, inRu, on, sales, type TableCase } from './cases.js'

export const local = true

export function buildTableCases(): TableCase[] {
  const rename = (id: string, about: string, name: string): TableCase => ({
    id,
    about,
    book: inRu(book(['S', { J1: '=SUM(Sales[Qty])' }, sales()])),
    steps: [{ op: { op: 'renameTable', table: 'Sales', name } }],
    read: on('S', 'A1:J5'),
  })
  return [
    rename('rename-bool', 'a table renamed ИСТИНА, a boolean of ru-RU', 'ИСТИНА'),
    rename('rename-bool-lower', 'a table renamed ложь, a boolean of ru-RU in another case', 'ложь'),
    rename('rename-true', 'a table renamed TRUE, which Excel ru-RU takes as a defined name', 'TRUE'),
  ]
}
