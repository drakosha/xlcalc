/**
 * Held out after `merge-third-held`, predicted before Excel ran it (`merge-top-held.predicted.json`): formulas, a name
 * and a formula on another sheet that read the empty top-left cell a merge moves a cell into, alone, in a range with
 * it and as an absolute reference. Cut and paste over a cell that formulas read makes their references `#REF!`.
 */
import type { StructureOp } from '../../../src/contract.js'
import type { StructCase, StructCases } from './cases.js'
import { merge } from './merge.js'
import type { StoredBook, StoredCell } from '../stored.js'
import { formulaCell as f } from '../stored.js'

/** S: an empty A1 before B1 and C1, read alone, in ranges and absolutely; an empty A5 read by a name and another sheet. */
function topBook(): StoredBook {
  return {
    sheets: [
      {
        name: 'S',
        cells: {
          B1: { value: 2 },
          C1: { value: 3 },
          H1: f('A1'),
          H2: f('$A$1+B1'),
          H3: f('SUM(A1:B1)'),
          H4: f('COUNTA(A1:A2)'),
          H5: f('IF(A1="",1,0)'),
          B5: { value: 5 },
          H6: f('top'),
        },
      },
      { name: 'T', cells: { A1: f('S!A5'), A2: f('S!A1+S!B1') } },
    ],
    names: [{ name: 'top', ref: "'S'!$A$5" }],
  }
}

export function buildStructCases(): StructCases {
  const cases: StructCase[] = []
  const add = (id: string, op: StructureOp) => cases.push({ id, book: 'top', op })
  add('merge-A1-C1-read-top', merge('A1:C1'))
  add('merge-A5-B5-name-top', merge('A5:B5'))
  add('merge-A1-B2-read-top', merge('A1:B2'))
  return { books: { top: topBook() }, cases }
}
