/**
 * The fill handle over hidden rows (Excel for Windows, `Range.AutoFill`), held out after the set `hidden` showed that the
 * handle fills rows hidden by hand and by a filter alike: text and a formula, numbers and texts filled down over them by
 * default and as copies, a fill up, and sources that hold hidden rows. Predicted by the engine before Excel ran it
 * (`fill-hidden-held.predicted.json`).
 */
import type { CellInput } from '../../../src/contract.js'
import { caseAdder, type StructCases } from './cases.js'
import { fill } from './fill.js'
import { FILES, HIDDEN, sheets } from './hidden.js'

/** F2:F3 the numbers 1 and 2, G2:G3 the texts x and y, H5:H8 texts and I5:I8 numbers over the hidden rows 7:8, J16 text. */
const EXTRA: Record<string, CellInput> = {
  F2: { value: 1 },
  F3: { value: 2 },
  G2: { value: 'x' },
  G3: { value: 'y' },
  ...Object.fromEntries(['p', 'q', 'r', 's'].map((v, i) => [`H${i + 5}`, { value: v }])),
  ...Object.fromEntries([1, 2, 3, 4].map((v, i) => [`I${i + 5}`, { value: v }])),
  J16: { value: 'up' },
}

export function buildStructCases(): StructCases {
  const books = { manual: { sheets: sheets(HIDDEN, false, EXTRA) }, filtered: { sheets: sheets(HIDDEN, true, EXTRA) } }
  const all = Object.keys(books).flatMap((book) => {
    const { cases, add } = caseAdder(book)
    add(`${book}-text-formula`, fill('D2:E2', 'D2:E16'))
    add(`${book}-text-formula-copy`, fill('D2:E2', 'D2:E16', 'copy'))
    add(`${book}-numbers`, fill('F2:F3', 'F2:F16'))
    add(`${book}-numbers-copy`, fill('F2:F3', 'F2:F16', 'copy'))
    add(`${book}-texts`, fill('G2:G3', 'G2:G16'))
    add(`${book}-up`, fill('J16', 'J2:J16'))
    add(`${book}-source-hidden-texts`, fill('H5:H8', 'H5:H16'))
    add(`${book}-source-hidden-numbers`, fill('I5:I8', 'I5:I16'))
    return cases
  })
  return { books, cases: all, files: FILES }
}
