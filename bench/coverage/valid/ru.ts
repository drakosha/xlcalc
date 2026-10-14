/**
 * Data validation in Excel ru-RU (`run.ts valid ru ru`), predicted before Excel ran it (`ru.predicted.json`): literal
 * lists as a file holds them, with commas and with semicolons; lists, bounds and formulas typed into the object model in
 * Russian (the list separator `;`, the decimal comma, local dates, Russian function names), and what Excel stores of them.
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import type { ValidCases } from './cases.js'
import { valueCell as v } from '../stored.js'

function book(): BookInput {
  const cells: Record<string, CellInput> = {
    A2: v('b'), A3: v('a,b'), A4: v('a;b'),
    B2: v('b'), B3: v('a;b;c'),
    C2: v('y'), C3: v('x;y'),
    D2: v('q'), D3: v('p,q'),
    E2: v(2), E3: v(3),
    F2: v(5), F3: v('x'),
    G2: v(45670), G3: v(45700),
  }
  return {
    locale: 'ru-RU',
    sheets: [
      {
        name: 'S',
        cells,
        validations: [
          { range: 'A2:A4', type: 'list', formula1: '"a,b,c"' },
          { range: 'B2:B3', type: 'list', formula1: '"a;b;c"' },
          { range: 'E2:E3', type: 'decimal', formula1: '1.5', formula2: '2.5' },
        ],
      },
    ],
  }
}

export function buildValidCases(): ValidCases {
  return {
    books: { ru: book() },
    cases: [
      { id: 'file', book: 'ru' },
      {
        id: 'typed',
        book: 'ru',
        make: [
          { sheet: 'S', range: 'C2:C3', rule: { type: 'list', formula1: '"x,y,z"' }, excel: ['x;y;z'] },
          { sheet: 'S', range: 'D2:D3', rule: { type: 'list', formula1: '"p,q"' }, excel: ['p,q'] },
          { sheet: 'S', range: 'F2:F3', rule: { type: 'custom', formula1: 'ISNUMBER(F2)' }, excel: ['=ЕЧИСЛО(F2)'] },
          { sheet: 'S', range: 'G2:G3', rule: { type: 'date', formula1: '45658', formula2: '45688' }, excel: ['01.01.2025', '31.01.2025'] },
        ],
      },
    ],
    window: 'A1:H5',
  }
}
