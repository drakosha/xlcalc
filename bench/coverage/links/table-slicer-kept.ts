/**
 * What of a table's slicer Excel opens and keeps from the file `writeKept` writes over a book Excel saved
 * (`test/excel/win/table-slicer/`), each case the file itself: its order and look changed, the table and its column renamed,
 * a column added, the slicer renamed. `SAVED` and `refused` were committed before Excel ran this set. The model: the part
 * of a cache in Excel's form, the ids in the file kept, so every case opens and saves with what it holds. And the file as
 * the writer gave it before, a cache in the form of a PivotTable's with no id, which Excel refuses.
 */
import { readFileSync } from 'node:fs'
import type { BookInput, StructureOp } from '../../../src/contract.js'
import { createSession } from '../../../src/index.js'
import { openKept, writeKept } from '../../../src/xlsx-keep.js'
import type { LinksCase, LinksCases } from './cases.js'
import { REFUSED, subs } from './file-refusals.js'

const CACHE = 'xl/slicerCaches/slicerCache1.xml'
const excel = (id: string) => new Uint8Array(readFileSync(new URL(`../../../test/excel/win/table-slicer/${id}.xlsx`, import.meta.url)))

/** The file over Excel's book, with the case's mirror sheets, after the operations and the cells typed. */
const keptFile = (from: string, ops: StructureOp[], typed: Record<string, string> = {}) => async (b: BookInput) => {
  const original = excel(from)
  const { book } = await openKept(original)
  const s = createSession({ ...book, sheets: [...book.sheets, ...b.sheets.filter((x) => /^W[TNX]$/.test(x.name))] })
  for (const op of ops) {
    const r = s.editStructure(op)
    if (!r.ok) throw new Error(r.message)
  }
  s.setCells(Object.entries(typed).map(([cell, value]) => ({ sheet: 'Data', cell, input: { value } })))
  const { data } = await writeKept(original, s.getSnapshot())
  s.close()
  return data
}

const SORT: StructureOp = { op: 'editSlicer', slicer: 'Product', cache: { sort: 'descending', customListSort: false, noData: { mark: false } } }
const base = (id: string, about: string, file: LinksCase['file'], e: Partial<LinksCase> = {}): LinksCase => ({
  id,
  about,
  sheets: [{ name: 'Data', cells: {} }],
  linked: {},
  links: [],
  read: [{ sheet: 'Data', range: 'C2' }],
  predicted: { 'Data!C2': 10 },
  written: true,
  file,
  ...e,
})

/** What the file Excel saved holds, predicted: the table's name, each column's id, the cache's source and what its table element says, slicers in the part, rows hidden. */
export interface Held {
  table: string
  columns: [number, string][]
  source: string
  cache: string
  slicers: string[]
  hidden: number
  filters?: string[]
}
const IDS: [number, string][] = [[2, 'Region'], [3, 'Product'], [4, 'Amount']]
const PLAIN: Held = { table: 'Sales', columns: IDS, source: 'Product', cache: 'tableId="1" column="3"', slicers: ['Product'], hidden: 0 }
const CHOSEN: Held = { ...PLAIN, hidden: 2, filters: ['Apple', 'Cherry'] }

export const SAVED: Record<string, Held> = {
  same: PLAIN,
  order: { ...PLAIN, cache: 'tableId="1" column="3" sortOrder="descending" customListSort="0" crossFilter="none"' },
  'cross-last': { ...PLAIN, cache: 'tableId="1" column="3" crossFilter="showItemsWithNoData"' },
  look: PLAIN,
  'table-renamed': { ...PLAIN, table: 'Goods' },
  'column-renamed': { ...PLAIN, columns: [[2, 'Region'], [3, 'Goods'], [4, 'Amount']], source: 'Goods' },
  'column-added': { ...PLAIN, columns: [[2, 'Region'], [5, 'Column1'], [3, 'Product'], [4, 'Amount']] },
  'slicer-renamed': { ...PLAIN, slicers: ['Goods'] },
  'chosen-same': CHOSEN,
  'chosen-order': { ...CHOSEN, cache: PLAIN.cache + ' sortOrder="descending"' },
}

export function buildLinksCases(): LinksCases {
  return {
    cases: [
      base('same', 'nothing changed but the mirror sheets', keptFile('plain', [])),
      base('order', 'its order and display changed', keptFile('plain', [SORT])),
      base('cross-last', 'its items with no data not last', keptFile('plain', [{ op: 'editSlicer', slicer: 'Product', cache: { noData: { last: false } } }])),
      base('look', 'its style changed', keptFile('plain', [{ op: 'editSlicer', slicer: 'Product', set: { style: 'SlicerStyleDark2' } }])),
      base('table-renamed', 'the table renamed', keptFile('plain', [{ op: 'renameTable', table: 'Sales', name: 'Goods' }])),
      base('column-renamed', 'its column renamed by its header', keptFile('plain', [], { B1: 'Goods' })),
      base('column-added', 'a column added to the table before its own', keptFile('plain', [{ op: 'insertColumns', sheet: 'Data', column: 'B', count: 1 }]), { read: [{ sheet: 'Data', range: 'D2' }], predicted: { 'Data!D2': 10 } }),
      base('slicer-renamed', 'the slicer renamed: its frame written again', keptFile('plain', [{ op: 'editSlicer', slicer: 'Product', set: { name: 'Goods' } }])),
      base('chosen-same', 'a file with a choice: the filter and the hidden rows stay', keptFile('chosen', [])),
      base('chosen-order', 'a file with a choice, its order changed', keptFile('chosen', [{ op: 'editSlicer', slicer: 'Product', cache: { sort: 'descending' } }])),
      base('tabular-no-id', 'the cache as 0.87.0 wrote it after such a change', keptFile('plain', []), {
        edit: subs([CACHE, /<extLst>.*<\/extLst>/s, '<pivotTables/><data><tabular pivotCacheId="undefined" sortOrder="descending"><items count="0"></items></tabular></data>']),
        ...REFUSED,
      }),
    ],
  }
}
