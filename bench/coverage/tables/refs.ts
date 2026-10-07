/**
 * Structured references computed: every part of a table, the formula's own row from inside and outside the table and
 * from another sheet, tables without a header row and with a totals row, escapes in column names, names over tables,
 * arrays inside and next to tables.
 */
import { book, on, sales, table, type TableCase } from './cases.js'

/** Formulas down column J of S, one per row from J1. */
const down = (fs: string[]) => Object.fromEntries(fs.map((f, i) => [`J${i + 1}`, `=${f}`]))

export function buildTableCases(): TableCase[] {
  const cases: TableCase[] = []
  const add = (c: TableCase) => cases.push(c)
  const parts = [
    'SUM(Sales[Qty])', 'Sales[@Qty]', 'SUM(Sales[[Qty]:[Price]])', 'ROWS(Sales[#All])', 'ROWS(Sales[#Data])', 'COLUMNS(Sales[#All])',
    'Sales[[#Headers],[Qty]]', 'ROWS(Sales[[#Headers],[#Data]])', 'ISERROR(Sales[#Totals])', 'ROWS(Sales)', 'ROWS(Sales[])', 'COUNTA(Sales[[#Headers],[#Data],[Region]])',
    'SUM(Sales[[#Data],[#Totals],[Qty]])', 'ISERROR(Sales[[#Totals],[Qty]])', 'SUM(Sales[[#Headers],[Qty]:[Price]])', 'SUM(Sales[[#All],[Qty]])', 'SUM(Sales[[Price]:[Qty]])',
    'SUM(Sales[[#Data],[Qty]:[Price]])', 'SUMIF(Sales[Region],"North",Sales[Total])', 'INDEX(Sales,2,2)', 'ROW(Sales[#Headers])', 'COLUMN(Sales[Price])',
    'SUM((Sales[Qty],Sales[Price]))', 'AREAS((Sales[Qty],Sales[Price]))', 'COUNTIFS(Sales[Region],"North",Sales[Qty],">2")', 'XLOOKUP("East",Sales[Region],Sales[Price])',
    'MAX(Sales[Total])', 'TYPE(Sales[Qty])', 'ISREF(Sales[Qty])', 'sales[qty]',
  ]
  add({ id: 'parts', about: 'every part of a table from a cell outside it, rows 1 to 30', book: book(['S', down(parts), sales()]), read: on('S', 'J1:J30') })

  const rows = {
    J2: '=Sales[@Qty]', J3: '=ROWS(Sales[@])', J4: '=Sales[@Price]+1', J5: '=ISREF(Sales[@Qty])', K2: '=SUM(Sales[@[Qty]:[Price]])', K3: '=Sales[@Total]*2',
    L3: '=Sales[@]', L4: '=Sales[[#This Row],[Region]:[Qty]]', J1: '=Sales[@Qty]', J7: '=Sales[@Qty]', K9: '=Sales[@Price]',
  }
  add({
    id: 'this-row',
    about: "the formula's row: in the data rows, above and below them, from another sheet",
    book: book(['S', rows, sales()], ['Other', { A3: '=Sales[@Qty]', B5: '=Sales[@Price]', C9: '=Sales[@Qty]', D1: '=Sales[@Region]', E4: '=SUM(Sales[@[Qty]:[Price]])' }]),
    read: [...on('S', 'J1:O9'), ...on('Other', 'A1:E9')],
  })

  add({
    id: 'inside',
    about: 'references written without a table name in its cells, and in the calculated column',
    book: book([
      'S',
      { F2: '=SUM(Sales[Qty])' },
      table('Sales', 'A1', ['Qty', 'Price', 'Total', 'Share', 'Run'], [
        [2, 10, '=[@Qty]*[@Price]', '=[@Total]/SUM([Total])', '=SUM([Qty])'],
        [3, 20, '=[@Qty]*[@Price]', '=[@Total]/SUM([Total])', '=ROWS([#Headers])'],
        [5, 30, '=[@Qty]*[@Price]', '=[@Total]/SUM([Total])', '=COUNTA(Sales[[#All],[Qty]])'],
      ]),
    ]),
    read: on('S', 'A1:F4'),
  })

  add({
    id: 'no-headers',
    about: 'a table without a header row: the header specifiers are #REF!, the data starts at its first row',
    book: book(['S', down(['ISERROR(T[#Headers])', 'ROWS(T[#All])', 'SUM(T[b])', 'ROWS(T[[#Headers],[#Data]])', 'T[@a]', 'ISERROR(T[[#Headers],[a]])']), table('T', 'A1', ['a', 'b'], [[1, 10], [2, 20], [3, 30]], { headers: false })]),
    read: on('S', 'J1:J6', 'A1:B4'),
  })

  add({
    id: 'totals',
    about: 'a totals row: its specifiers, SUBTOTAL over the column, a label, and the formula row in it',
    book: book([
      'S',
      down(['SUM(T[[#Totals],[b]])', 'ROWS(T)', 'ROWS(T[#All])', 'SUM(T[[#Data],[#Totals],[b]])', 'T[[#Totals],[a]]', 'SUM(T[b])', 'ROWS(T[[#Headers],[#Data]])']),
      table('T', 'A1', ['a', 'b', 'c'], [[1, 10, '=[@b]*2'], [2, 20, '=[@b]*2'], [3, 30, '=[@b]*2']], { totals: ['Total', '=SUBTOTAL(109,[b])', '=[@b]'] }),
    ]),
    read: on('S', 'J1:J7', 'A1:C5'),
  })

  const wierd = table('Odd', 'A1', ['a b', 'x#y', "p'q", '[br]', 'Total $', '@at', '2024', 'Ünï', 'Сумма'], [[1, 2, 3, 4, 5, 6, 7, 8, 9], [10, 20, 30, 40, 50, 60, 70, 80, 90]])
  add({
    id: 'escapes',
    about: "column names with spaces and the characters Excel escapes with '",
    book: book(['S', down(['SUM(Odd[a b])', 'SUM(Odd[[a b]])', "SUM(Odd[x'#y])", "SUM(Odd[p''q])", "SUM(Odd['[br']])", 'SUM(Odd[Total $])', "SUM(Odd['@at])", 'SUM(Odd[2024])', 'SUM(Odd[Ünï])', 'SUM(Odd[Сумма])', "SUM(Odd[[x'#y]:[p''q]])", "Odd[[#Headers],['@at]]", 'Odd[[#Headers],[2024]]']), wierd]),
    read: on('S', 'J1:J13'),
  })

  const nm = book(['S', down(['SUM(nm)', 'COUNTA(nh)', 'nr', 'nr*2']), sales()], ['Other', { A2: '=nr', A4: '=nr', A9: '=nr', B1: '=SUM(nm)' }])
  nm.names = [
    { name: 'nm', ref: 'Sales[Qty]' },
    { name: 'nh', ref: 'Sales[#Headers]' },
    { name: 'nr', ref: 'Sales[[#This Row],[Price]]' },
  ]
  add({ id: 'names', about: "names over a table, the formula's row of the cell that uses one", book: nm, read: [...on('S', 'J1:J4'), ...on('Other', 'A1:B9')] })

  const sp = book(['S', { F1: '=SEQUENCE(3)', H6: '=SEQUENCE(1,3)', J1: '=Sales[a]', K1: '=Sales[#Headers]', J7: '=Sales[[#All],[b]]' }, table('Sales', 'A3', ['a', 'b'], [[1, '=SEQUENCE(2)'], [2, '=[@a]*2']]), table('Side', 'F3', ['h'], [[1], [2]])])
  add({ id: 'spill', about: 'arrays inside a table, into one from above, and references to tables that spill', book: sp, read: on('S', 'A1:N12') })
  return cases
}
