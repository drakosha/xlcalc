/**
 * Held out: tables named by the rules of names, where they were measured for defined names (`names-more-held`): a
 * question mark and a backslash inside a table's name, a question mark first, in en-US; predicted by the engine before
 * Excel ran them (`names-tables-held.predicted.json`). Excel's object model makes a name of one it does not take.
 */
import type { StructureOp } from '../../../src/contract.js'
import { book, type CellSpec, on, sales, table, type TableCase, type TableStep, typedFormula } from './cases.js'

const op = (o: StructureOp): TableStep => ({ op: o })
const OVER: Record<string, CellSpec> = { J1: '=SUM(Sales[Qty])', J2: '=ROWS(Sales)' }

export function buildTableCases(): TableCase[] {
  const cases: TableCase[] = []
  const rename = (id: string, about: string, name: string, then: TableStep[] = []) =>
    cases.push({ id, about, book: book(['S', OVER, sales()]), steps: [op({ op: 'renameTable', table: 'Sales', name }), ...then], read: on('S', 'A1:K5') })
  rename('rename-question', 'a table renamed with a question mark inside (Прод?ажи), then a reference to it typed', 'Прод?ажи', [typedFormula('J3', 'SUM(Прод?ажи[Price])')])
  rename('rename-backslash', 'a table renamed with a backslash inside (Прод\\ажи), then a reference to it typed', 'Прод\\ажи', [typedFormula('J3', 'SUM(Прод\\ажи[Price])')])
  rename('rename-question-first', 'a table renamed with a question mark first (?Продажи): no name, Excel makes one of it', '?Продажи')
  const t = table('Прод?ажи', 'A1', ['Region', 'Qty'], [['North', 2], ['South', 3]])
  cases.push({ id: 'file-question', about: 'a file holding a table named with a question mark inside, and a formula over it', book: book(['S', { J1: '=SUM(Прод?ажи[Qty])' }, t]), read: on('S', 'A1:J3') })
  return cases
}
