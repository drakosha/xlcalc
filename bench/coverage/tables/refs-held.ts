/**
 * Held out: structured references the set `refs` does not have, predicted by the engine before Excel computed them
 * (`refs-held.predicted.json`, from `predict.ts`). A table away from A1 with a totals row, the formula's row at the
 * table's edges, more escapes, names, arrays beside tables.
 */
import { book, on, table, type TableCase } from './cases.js'

export function buildTableCases(): TableCase[] {
  const orders = () =>
    table('Orders', 'C4', ['Code', 'Qty', 'Amount'], [['a', 1, 10.5], ['b', 4, 20], ['c', 2, -3], ['b', 7, 0.25]], { totals: ['Total', '=SUBTOTAL(109,[Qty])', '=SUBTOTAL(101,[Amount])'] })
  const formulas = {
    H1: '=AVERAGE(Orders[Amount])', H2: '=SUMPRODUCT(Orders[Qty],Orders[Amount])', H3: '=MATCH("b",Orders[Code],0)', H4: '=Orders[@Code]', H5: '=Orders[@Code]',
    H6: '=Orders[@Qty]*2', H8: '=Orders[@Amount]', H9: '=Orders[@Amount]', H10: '=Orders[@Code]', H11: '=ROWS(Orders[[#Data],[#Totals]])', H12: '=INDEX(Orders[#All],1,2)',
    H13: '=Orders[[#Totals],[Amount]]', H14: '=COUNTA(Orders[[#All],[Code]])', H15: '=SUM(Orders[[#Headers],[#Data],[Qty]:[Amount]])', H16: '=COLUMNS(Orders[[Code]:[Amount]])',
    H17: '=XLOOKUP(7,Orders[Qty],Orders[Code])', H18: '=MAX(Orders[[#Totals],[Qty]:[Amount]])', H19: '=SUMIFS(Orders[Amount],Orders[Code],"b",Orders[Qty],">5")',
  }
  const cases: TableCase[] = [
    {
      id: 'offset',
      about: 'a table at C4 with a totals row, read from beside it: its rows at the edges of the data, the totals',
      book: book(['S', formulas, orders()], ['T', { A5: '=Orders[@Qty]', A9: '=Orders[@Qty]', B4: '=Orders[@Code]', C10: '=Orders[[#Totals],[Qty]]' }]),
      read: [...on('S', 'C4:F9', 'H1:H19'), ...on('T', 'A1:C10')],
    },
    {
      id: 'totals-row',
      about: "formulas in a totals row: the formula's row there, the totals of another column, the data above",
      book: book([
        'S',
        {},
        table('Tab', 'A1', ['x', 'y', 'z'], [[1, 2, '=[@x]+[@y]'], [3, 4, '=[@x]+[@y]']], { totals: ['=ROWS([#Data])', '=[@y]', '=Tab[[#Totals],[x]]+SUM([z])'] }),
      ]),
      read: on('S', 'A1:C4'),
    },
    {
      id: 'escapes',
      about: "more names that take ' before a character: a # first, brackets inside, two apostrophes",
      book: book([
        'S',
        { G1: "=SUM(E[['#n]:[a'[1']]])", G2: "=E[[#Headers],[it''s]]", G3: "=SUM(E['#n])+SUM(E[a'[1']])", G4: "=E[@it''s]", G5: '=COUNT(E[#Data])' },
        table('E', 'A1', ['#n', "a[1]", "it's", 'Sp ace'], [[1, 2, 3, 4], [5, 6, 7, 8], [9, 10, 11, 12]]),
      ]),
      read: on('S', 'G1:G5'),
    },
    {
      id: 'names',
      about: 'names: a bare table, its formula row used on rows inside and outside the data, a name over another sheet\'s table',
      book: {
        ...book(['S', { E1: '=ROWS(nt)', E2: '=ny', E3: '=ny', E5: '=ny', F1: '=SUM(nt)' }, table('Tab', 'A1', ['x', 'y'], [[1, 2], [3, 4], [5, 6]])], ['U', { A3: '=ny*10', A4: '=ny*10', B1: '=COUNT(nall)' }]),
        names: [
          { name: 'nt', ref: 'Tab' },
          { name: 'ny', ref: 'Tab[[#This Row],[y]]' },
          { name: 'nall', ref: 'Tab[#All]' },
        ],
      },
      read: [...on('S', 'E1:F5'), ...on('U', 'A1:B4')],
    },
    {
      id: 'arrays',
      about: 'an array from the left into a table, a one-value array inside one, a reference to a table spilled below it',
      book: book(['S', { A5: '=SEQUENCE(1,3)', A1: '=SEQUENCE(1,1)', A9: '=Tab[#All]', E1: '=TRANSPOSE(Tab[y])' }, table('Tab', 'B4', ['x', 'y'], [[1, '={7}'], [2, '=SUM(SEQUENCE(3))']])]),
      read: on('S', 'A1:H12'),
    },
  ]
  return cases
}
