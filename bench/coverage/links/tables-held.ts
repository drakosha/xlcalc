/**
 * Structured references to a table of a linked book, held out: a host's books hold `[1]Таблица[[#This Row],[Столбец]]`,
 * which the engine did not parse, and Excel stored `#REF!` for it. The table's own row and its column, a whole column in
 * SUM, `#All`, the form with `!` after the link, a Latin name, and a link whose book defines a name like the table.
 *
 * `predicted` was committed before Excel ran the set. The guess: Excel takes each form (no repair window), and a
 * reference to a table of a closed book is `#REF!`, whatever the link stores; the alternatives are the stored values
 * of the table's cells, read through the sheet the link names.
 */
import type { LinkedBook, LinksCase, LinksCases } from './cases.js'

const REF = { error: 'REF' } as const

/** A book with a table over A1:B4 of its sheet, the header row and three rows, as a link stores the cells read. */
const book = (names?: LinkedBook['names']): LinkedBook => ({
  target: 'Book2.xlsx',
  sheets: [{ name: 'Лист1', cells: { A1: 'Столбец', B1: 'Другой', A2: 1, B2: 10, A3: 2, B3: 20, A4: 3, B4: 30 } }],
  ...(names ? { names } : {}),
})

const rows = (f: (r: number) => string) => ({ A2: f(2), A3: f(3), A4: f(4) })

const CASES: LinksCase[] = [
  {
    id: 'this-row',
    about: 'the row of the formula in a column of a table of a linked book, rows 2 to 4',
    sheets: [{ name: 'S', cells: {} }],
    linked: { S: rows(() => '[1]Таблица[[#This Row],[Столбец]]') },
    links: [book()],
    read: [{ sheet: 'S', range: 'A2:A4' }],
    predicted: { 'S!A2': REF, 'S!A3': REF, 'S!A4': REF },
    alternatives: { 'S!A2': [1], 'S!A3': [2], 'S!A4': [3] },
  },
  {
    id: 'column',
    about: 'a whole column, #All, the row with an operator, a Latin table name',
    sheets: [{ name: 'S', cells: {} }],
    linked: { S: { A1: 'SUM([1]Таблица[Столбец])', A2: 'ROWS([1]Таблица[#All])', A3: '[1]Таблица[[#This Row],[Другой]]*2', A4: 'IFERROR([1]Table1[[#This Row],[Col]],"e")' } },
    links: [book()],
    read: [{ sheet: 'S', range: 'A1:A4' }],
    predicted: { 'S!A1': REF, 'S!A2': REF, 'S!A3': REF, 'S!A4': 'e' },
    alternatives: { 'S!A1': [6], 'S!A2': [4], 'S!A3': [40] },
  },
  {
    id: 'bang',
    about: 'the form with ! after the link',
    sheets: [{ name: 'S', cells: {} }],
    linked: { S: rows(() => '[1]!Таблица[[#This Row],[Столбец]]') },
    links: [book()],
    read: [{ sheet: 'S', range: 'A2:A4' }],
    predicted: { 'S!A2': REF, 'S!A3': REF, 'S!A4': REF },
    alternatives: { 'S!A2': [1], 'S!A3': [2], 'S!A4': [3] },
  },
  {
    id: 'named',
    about: 'the linked book defines a name like the table over its rows',
    sheets: [{ name: 'S', cells: {} }],
    linked: { S: { ...rows(() => '[1]Таблица[[#This Row],[Столбец]]'), A5: 'SUM([1]Таблица[Столбец])' } },
    links: [book([{ name: 'Таблица', ref: '=Лист1!$A$2:$B$4' }])],
    read: [{ sheet: 'S', range: 'A2:A5' }],
    predicted: { 'S!A2': REF, 'S!A3': REF, 'S!A4': REF, 'S!A5': REF },
    alternatives: { 'S!A2': [1], 'S!A3': [2], 'S!A4': [3], 'S!A5': [6] },
  },
]

export const buildLinksCases = (): LinksCases => ({ cases: CASES })
