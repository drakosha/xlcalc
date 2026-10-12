/**
 * Links written by `writeXlsx`: the cases of `basics-held` but `inserted`, each file the writer writes of the case's
 * book with its links as `externalBooks` (`written`), and what the files made by hand had not: cells stored empty
 * (`<cell r="A4"/>`) and empty text, on a sheet with the mark of a failed update and on one without; the errors of Excel
 * 365 among the stored values; a link without sheets and with an empty target (the place the reader keeps for a DDE or
 * OLE link) between two others.
 *
 * `predicted` was committed before Excel ran the set. The guess: Excel opens every file without the repair window and
 * shows for a case of `basics-held` what it showed for that file made by hand (the rules of the measurements: a cell
 * the link holds no value of is blank, and `#REF!` once an edit entered a formula over the link and marked its sheets;
 * a sheet without data is `#REF!` throughout; a name of a constant is `#REF!`); a cell stored empty is blank, with the
 * mark too, and empty text is text; `#SPILL!` and `#CALC!` stored come back as those errors; the link without sheets
 * keeps the numbers of the others.
 */
import { buildLinksCases as basicsHeld } from './basics-held.js'
import type { CachedValue, LinksCase, LinksCases } from './cases.js'

const REF = { error: 'REF' } as const

/** What Excel showed for the files made by hand where `basics-held` guessed otherwise, the same on the Mac and on Windows. */
const MEASURED: Record<string, Record<string, CachedValue>> = {
  // the fill downs and the copy marked every sheet of the link: the cells it holds no value of are #REF!
  cache: { 'S!A3': REF, 'S!A4': REF, 'S!A5': REF, 'S!A6': false, 'S!A15': REF, 'S!A24': REF, 'S!A28': 8, 'S!A35': REF, 'S!D10': REF },
  'refresh-error': { 'S!A2': REF, 'S!A3': false },
}

const HELD: LinksCase[] = basicsHeld()
  .cases.filter((c) => c.id !== 'inserted')
  .map(({ alternatives: _, ...c }) => ({ ...c, written: true, predicted: { ...c.predicted, ...MEASURED[c.id] } }))

const NEW: LinksCase[] = [
  {
    id: 'stored-empty',
    about: 'cells stored empty and empty text, on a sheet with the mark of a failed update and on one without',
    sheets: [{ name: 'S', cells: {} }],
    linked: {
      S: {
        A1: '[1]Sheet1!A2',
        A2: 'ISBLANK([1]Sheet1!A2)',
        A3: '[1]Sheet1!A2&"|"',
        A4: 'ISBLANK([1]Sheet1!B1)',
        A5: 'LEN([1]Sheet1!B1)',
        A6: '[1]Marked!A2',
        A7: 'ISBLANK([1]Marked!A2)',
        A8: '[1]Marked!A3',
        A9: 'COUNTA([1]Marked!A1:A3)',
      },
    },
    links: [
      {
        target: 'Book2.xlsx',
        sheets: [
          { name: 'Sheet1', cells: { A1: 1, A2: null, B1: '' } },
          { name: 'Marked', cells: { A1: 5, A2: null }, refreshError: true },
        ],
      },
    ],
    read: [{ sheet: 'S', range: 'A1:A9' }],
    written: true,
    predicted: { 'S!A1': 0, 'S!A2': true, 'S!A3': '|', 'S!A4': false, 'S!A5': 0, 'S!A6': 0, 'S!A7': true, 'S!A8': REF, 'S!A9': 2 },
    alternatives: { 'S!A4': [true], 'S!A6': [REF], 'S!A7': [false], 'S!A9': [3] },
  },
  {
    id: 'errors',
    about: "every error Excel stores of another book's cell, #SPILL! and #CALC! among them",
    sheets: [{ name: 'S', cells: {} }],
    linked: { S: { A1: '[1]Sheet1!A1', A2: '[1]Sheet1!A2', A3: 'ERROR.TYPE([1]Sheet1!A1)', A4: 'ERROR.TYPE([1]Sheet1!A2)', A5: '[1]Sheet1!A3', A6: '[1]Sheet1!A4', A7: '[1]Sheet1!A5', A8: '[1]Sheet1!A6' } },
    links: [{ target: 'Book2.xlsx', sheets: [{ name: 'Sheet1', cells: { A1: { error: 'SPILL' }, A2: { error: 'CALC' }, A3: { error: 'NULL' }, A4: { error: 'NUM' }, A5: { error: 'NAME' }, A6: { error: 'REF' } } }] }],
    read: [{ sheet: 'S', range: 'A1:A8' }],
    written: true,
    predicted: { 'S!A1': { error: 'SPILL' }, 'S!A2': { error: 'CALC' }, 'S!A3': 9, 'S!A4': 14, 'S!A5': { error: 'NULL' }, 'S!A6': { error: 'NUM' }, 'S!A7': { error: 'NAME' }, 'S!A8': REF },
    alternatives: { 'S!A1': [{ error: 'VALUE' }], 'S!A2': [{ error: 'VALUE' }], 'S!A3': [3], 'S!A4': [3] },
  },
  {
    id: 'placeholder',
    about: 'a link without sheets and with an empty target, the place of a DDE or OLE link, between two others',
    sheets: [{ name: 'S', cells: {} }],
    linked: { S: { A1: '[1]Sheet1!A1+[3]Data!A1', A2: '[3]Data!A1' } },
    links: [
      { target: 'Book2.xlsx', sheets: [{ name: 'Sheet1', cells: { A1: 5 } }] },
      { target: '', sheets: [] },
      { target: 'Data.xlsx', sheets: [{ name: 'Data', cells: { A1: 7 } }] },
    ],
    read: [{ sheet: 'S', range: 'A1:A2' }],
    written: true,
    predicted: { 'S!A1': 12, 'S!A2': 7 },
  },
]

/**
 * The predictions Excel did not confirm, the same on the Mac and on Windows: a link's stored #SPILL! and #CALC! show
 * `#VALUE!`, the alternative the case gave. The first file of the case held their text, which both Excels repair; the
 * writer then stored them as `#VALUE!` with their rich values, which Excel for Mac showed and saved as plain `#VALUE!`.
 */
export const MISSED: readonly string[] = ['errors S!A1', 'errors S!A2', 'errors S!A3', 'errors S!A4']

export function buildLinksCases(): LinksCases {
  return { cases: [...HELD, ...NEW] }
}
