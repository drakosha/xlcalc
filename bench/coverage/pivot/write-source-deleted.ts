/**
 * The set `write-source-deleted`: the book of `basic/amount-sum` with its source sheet `D` deleted, written by
 * `writeXlsx` (the source a link to a file by its path, as Excel writes it), opened by Excel and read, then refreshed.
 * Its file first:
 *
 *   npx tsx bench/coverage/pivot/write-source-deleted.ts
 */
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { editStructure } from '../../../src/index.js'
import { readXlsx, writeXlsx } from '../../../src/xlsx.js'
import type { PivotCase } from './cases.js'

const FILE = 'bench/windows/out/pivot-ours/source-deleted/link.xlsx'

export const buildPivotCases = (): PivotCase[] => [{ id: 'link', file: FILE, ops: [{ read: true }, { refresh: true }] }]

async function main() {
  const { book } = await readXlsx(readFileSync('test/excel/win/pivot/basic/amount-sum.xlsx'))
  const r = editStructure(book, { op: 'deleteSheet', sheet: 'D' })
  if (!r.ok) throw new Error(r.message)
  mkdirSync(FILE.slice(0, FILE.lastIndexOf('/')), { recursive: true })
  writeFileSync(FILE, (await writeXlsx({ book: r.book }, { values: 'omit' })).data)
}

if (process.argv[1]?.endsWith('write-source-deleted.ts')) await main()
