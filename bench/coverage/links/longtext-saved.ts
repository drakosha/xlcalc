/**
 * The file Excel for Mac saved of a text literal of 300 characters, as it is (`test/excel/links/longtext-file/literal-300.xlsx`, which holds
 * `LEN(_xlfn._LONGTEXT(…))` with the underscore), for Excel for Windows: the set `longtext-file` tried `_xlfn.LONGTEXT` alone, without it.
 *
 * Predicted before Excel ran it, from the function Excel for Windows has by the name `_LONGTEXT`: it opens the file and answers 300, with
 * the literal as Excel for Mac reads it back, and saves the same form. The reader takes no decision from this.
 */
import { readFileSync } from 'node:fs'
import type { LinksCases } from './cases.js'

const SAVED = new URL('../../../test/excel/links/longtext-file/literal-300.xlsx', import.meta.url)

export function buildLinksCases(): LinksCases {
  return {
    cases: [
      {
        id: 'mac-300',
        about: 'the file Excel for Mac saved of a text of 300',
        sheets: [{ name: 'S', cells: {} }],
        linked: {},
        links: [],
        read: [{ sheet: 'S', range: 'A1' }],
        predicted: { 'S!A1': 300 },
        written: true,
        file: async () => new Uint8Array(readFileSync(SAVED)),
      },
    ],
  }
}
