/**
 * The Sort command against Excel for Mac en-US: the book of `sort/<set>.ts` (`sortBook`), each case's column sorted by
 * itself without a header (`excel-sort.applescript`), and the column's formulas read back into
 * `test/excel/sort/<set>.json`; Windows' are `test/excel/win/sort/<set>.json` (`bench/windows/sort.ts`).
 * `test/sort-command-excel.test.ts` compares `valueComparator` with them without Excel.
 */
import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { writeXlsx } from '../../src/xlsx.js'
import { ensureExcel, excelVersion, runScriptWatched } from './excel.js'
import { ROOT, HERE } from './paths.js'
import { secondsSince, setModule, writeAnswers } from './stored.js'
import { rangeOf, sortBook, type SortSet } from './sort/cases.js'

const SCRIPT = join(HERE, 'excel-sort.applescript')

export async function stepSort(name: string | undefined): Promise<void> {
  const { set } = await setModule<{ set: SortSet }>('sort', name)
  await ensureExcel('en-US')
  const book = join(mkdtempSync(join(tmpdir(), 'fsort-')), 'book.xlsx')
  writeFileSync(book, (await writeXlsx({ book: sortBook(set) }, { values: 'omit' })).data)
  const t0 = Date.now()
  const specs = set.cases.map((c, i) => `${rangeOf(set, i).replace(':', ' ')} ${c.descending ? 'd' : 'a'}`)
  const got = await runScriptWatched(SCRIPT, book, set.cases.length, [specs.join('\n')])
  if (got === 'rejected') throw new Error('Excel did not open the book whole')
  const cols = JSON.parse(got) as unknown[][][]
  const answers = Object.fromEntries(set.cases.map((c, i) => [c.id, cols[i].map(([f]) => String(f))]))
  writeAnswers(join(ROOT, 'test/excel/sort', `${name}.json`), { meta: { version: excelVersion(), platform: 'Mac', locale: 'en-US', seconds: secondsSince(t0) }, answers })
  console.log(`sort ${name}: ${set.cases.length} cases`)
}
