/**
 * Tables in Excel ru-RU (`run.ts tables locale ru`): structured references as Excel shows them in Russian (`#Данные`,
 * `;` between items, `@` for the formula's row), formulas typed in Russian, and the names Excel gives new columns.
 */
import { inRu, on, sales, table, type TableCase } from './cases.js'

export const local = true

export function buildTableCases(): TableCase[] {
  const s = sales()
  const shown = {
    J1: '=SUM(Sales[Qty])', J2: '=ROWS(Sales[#All])', J3: '=ROWS(Sales[#Data])', J4: '=Sales[[#Headers],[Qty]]', J5: '=ISERROR(Sales[#Totals])',
    J6: '=SUM(Sales[[#Data],[#Totals],[Qty]])', J7: '=ROWS(Sales[[#Headers],[#Data]])', J8: '=SUM(Sales[[#All],[Qty]:[Price]])', J9: '=ROWS(Sales)',
    K2: '=Sales[@Qty]', K3: '=SUM(Sales[@[Qty]:[Price]])', K4: '=COUNTA(Sales[#This Row])', K5: '=Sales[[#This Row],[Region]]&"!"',
  }
  const totals = table('T', 'A1', ['a', 'b'], [[1, '=[@a]*2'], [2, '=[@a]*2']], { totals: ['=ROWS([#Data])', '=SUBTOTAL(109,[b])'] })
  return [
    { id: 'shown', about: 'every specifier and the formula row as Excel ru-RU shows them, outside and inside the table', book: inRu({ sheets: [{ name: 'S', cells: { ...s.cells, ...cellsOf(shown) }, tables: [s.table] }] }), read: on('S', 'A1:K9') },
    { id: 'totals', about: 'a totals row and a calculated column in ru-RU', book: inRu({ sheets: [{ name: 'S', cells: totals.cells, tables: [totals.table] }] }), read: on('S', 'A1:B4') },
    {
      id: 'typed',
      about: 'formulas typed in Russian: the specifiers in any case, the formula row as #Эта строка',
      book: inRu({ sheets: [{ name: 'S', cells: s.cells, tables: [s.table] }] }),
      steps: [
        { edit: { sheet: 'S', cell: 'J1', value: { local: '=СУММ(Sales[[#Заголовки];[Qty]])' } } },
        { edit: { sheet: 'S', cell: 'J2', value: { local: '=ЧСТРОК(Sales[#Все])' } } },
        { edit: { sheet: 'S', cell: 'J3', value: { local: '=Sales[[#Эта строка];[Qty]]' } } },
        { edit: { sheet: 'S', cell: 'J4', value: { local: '=СУММ(Sales[[#данные];[qty]])' } } },
        { edit: { sheet: 'S', cell: 'J5', value: { local: '=СУММ(Sales[[#Данные]; [Price]])' } } },
        { edit: { sheet: 'S', cell: 'J6', value: { local: '=ЧСТРОК(Sales[[#Заголовки];[#Данные]])' } } },
      ],
      read: on('S', 'J1:J6'),
    },
    {
      id: 'columns',
      about: 'a column inserted into a table and a header cleared: Excel ru-RU names them Столбец1 and Столбец2',
      book: inRu({ sheets: [{ name: 'S', cells: { ...s.cells, J1: { formula: 'SUM(Sales[[Region]:[Price]])' } }, tables: [s.table] }] }),
      steps: [{ op: { op: 'insertColumns', sheet: 'S', column: 'B' } }, { edit: { sheet: 'S', cell: 'A1', value: null } }],
      read: on('S', 'A1:K5'),
    },
  ]
}

function cellsOf(x: Record<string, string>) {
  return Object.fromEntries(Object.entries(x).map(([a, f]) => [a, { formula: f.slice(1) }]))
}
