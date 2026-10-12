/**
 * The first link: `=[1]Sheet1!A1+1` over another book's cell A1 stored as 5. It shows how Excel opens a file with a
 * link without updating it, on the Mac and on Windows, and how it shows such a formula.
 */
import type { LinksCases } from './cases.js'

export function buildLinksCases(): LinksCases {
  return {
    cases: [
      {
        id: 'one-cell',
        about: '[1]Sheet1!A1+1 over A1 = 5 stored in the link',
        sheets: [{ name: 'S', cells: { B1: { value: 1 } } }],
        linked: { S: { A1: '[1]Sheet1!A1+1' } },
        links: [{ target: 'Book2.xlsx', sheets: [{ name: 'Sheet1', cells: { A1: 5 } }] }],
        read: [{ sheet: 'S', range: 'A1:B1' }],
      },
    ],
  }
}
