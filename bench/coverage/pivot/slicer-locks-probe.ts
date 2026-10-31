/**
 * The set `slicer-locks-probe`: which part of the file `slicer-locks/xlsx` Excel does not open. The book of
 * `slicer-objects/place-mixed` (two slicers, two timelines) written by `writeXlsx` as read (`as-read`), with the locks
 * alone (`locks`), with the rename alone (`rename`), without its second timeline (`one-timeline`), without its timelines
 * (`no-timelines`), and `slicer/timeline-and-slicer` as read (`one-each`); into `bench/windows/out/pivot-ours/locks/`.
 * Excel opens each and reads its slicers' names. Predicted before Excel ran it: every file opens
 * (`slicer-locks-probe.predicted.json`). Write the files: `npx tsx bench/coverage/pivot/slicer-locks-probe.ts`.
 */
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { createSession, registerPivot, type BookInput, type StructureOp } from '../../../src/index.js'
import { pivotEngine } from '../../../src/pivot/index.js'
import { readXlsx, writeXlsx } from '../../../src/xlsx.js'
import type { PivotCase } from './cases.js'

const DIR = 'bench/windows/out/pivot-ours/locks'
const IDS = ['as-read', 'locks', 'rename', 'one-timeline', 'no-timelines', 'one-each']

export const buildPivotCases = (): PivotCase[] => IDS.map((id) => ({ id, file: `${DIR}/probe-${id}.xlsx`, ops: [{ slicerGet: id === 'one-each' ? 'Slicer_Product' : 'Slicer_Region', get: ['Name'] }] }))

const LOCKS: StructureOp[] = [
  { op: 'editSlicer', slicer: 'Slicer_Region', set: { locked: false, positionLocked: true } },
  { op: 'editTimeline', timeline: 'Timeline_Date', set: { locked: false } },
]
const RENAME: StructureOp[] = [{ op: 'editSlicer', slicer: 'Slicer_Product', set: { name: 'Products' } }]

/** The book with `ops` played in a session. */
function played(book: BookInput, ops: StructureOp[]): BookInput {
  const s = createSession(book)
  for (const op of ops) {
    const r = s.editStructure(op)
    if (!r.ok) throw new Error(r.message)
  }
  return s.getBook()
}

/** The book without the timelines `names` and their caches. */
const without = (book: BookInput, names: string[]): BookInput => {
  const caches = book.timelineCaches?.filter((c) => !names.includes(c.name))
  const out: BookInput = { ...book, sheets: book.sheets.map((s) => (s.timelines ? { ...s, timelines: s.timelines.filter((t) => !names.includes(t.name)) } : s)) }
  if (caches?.length) out.timelineCaches = caches
  else delete out.timelineCaches
  return out
}

if (import.meta.url === `file://${process.argv[1]}`) {
  registerPivot(pivotEngine)
  mkdirSync(DIR, { recursive: true })
  const mixed = (await readXlsx(readFileSync('test/excel/win/pivot/slicer-objects/place-mixed.xlsx'))).book
  const each = (await readXlsx(readFileSync('test/excel/win/pivot/slicer/timeline-and-slicer.xlsx'))).book
  const books: Record<string, BookInput> = {
    'as-read': mixed,
    locks: played(mixed, LOCKS),
    rename: played(mixed, RENAME),
    'one-timeline': without(mixed, ['Timeline_When']),
    'no-timelines': without(mixed, ['Timeline_Date', 'Timeline_When']),
    'one-each': each,
  }
  for (const id of IDS) writeFileSync(`${DIR}/probe-${id}.xlsx`, (await writeXlsx({ book: books[id] })).data)
  console.log('written', IDS.length)
}
