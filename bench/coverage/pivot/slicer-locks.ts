/**
 * The set `slicer-locks`: the engine's files with slicers and timelines locked and unlocked opened in Excel for Windows.
 * Over the book of `slicer-objects/place-mixed` (two slicers, two timelines), a session sets Slicer_Region's Locked off
 * and its Disable resizing and moving on, Timeline_Date's Locked off, and renames Slicer_Product, its locks as they
 * were (both by default); the book written by `writeXlsx` (`xlsx`) and by `writeKept` over the file (`kept`), into
 * `bench/windows/out/pivot-ours/locks/`. Excel opens each, reads each one's `Locked` and `DisableMoveResizeUI`, and
 * saves it. And a slicer named a space through `Slicers.Add` (`space`). Predicted before Excel ran it:
 * `slicer-locks.predicted.json`, each read as written. Write the files: `npx tsx bench/coverage/pivot/slicer-locks.ts`.
 */
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { createSession, registerPivot } from '../../../src/index.js'
import { pivotEngine } from '../../../src/pivot/index.js'
import { openKept, writeKept } from '../../../src/xlsx-keep.js'
import { writeXlsx } from '../../../src/xlsx.js'
import type { PivotCase } from './cases.js'

const DIR = 'bench/windows/out/pivot-ours/locks'
const GET = ['Name', 'Locked', 'DisableMoveResizeUI']
const NAMES = ['Slicer_Region', 'Products', 'Timeline_Date', 'Timeline_When']

export function buildPivotCases(): PivotCase[] {
  const opened = (how: string): PivotCase => ({ id: how, file: `${DIR}/${how}.xlsx`, ops: NAMES.map((n) => ({ slicerGet: n, get: GET })) })
  return [opened('xlsx'), opened('kept'), { id: 'space', ops: [{ rows: ['Region'] }, { value: { field: 'Price' } }, { slicer: 'Region', name: 'Slicer_Region', slicerName: ' ' }, { slicerGet: ' ', get: GET }] }]
}

/** The session's book over `place-mixed` with the locks set and the rename made, and the file it was read from. */
async function locked(): Promise<{ data: Uint8Array; session: ReturnType<typeof createSession> }> {
  const data = new Uint8Array(readFileSync('test/excel/win/pivot/slicer-objects/place-mixed.xlsx'))
  const s = createSession((await openKept(data)).book)
  for (const op of [
    { op: 'editSlicer', slicer: 'Slicer_Region', set: { locked: false, positionLocked: true } },
    { op: 'editTimeline', timeline: 'Timeline_Date', set: { locked: false } },
    { op: 'editSlicer', slicer: 'Slicer_Product', set: { name: 'Products' } },
  ] as const) {
    const r = s.editStructure(op)
    if (!r.ok) throw new Error(r.message)
  }
  return { data, session: s }
}

if (import.meta.url === `file://${process.argv[1]}`) {
  registerPivot(pivotEngine)
  mkdirSync(DIR, { recursive: true })
  const { data, session } = await locked()
  writeFileSync(`${DIR}/xlsx.xlsx`, (await writeXlsx({ book: session.getBook() })).data)
  writeFileSync(`${DIR}/kept.xlsx`, (await writeKept(data, session.getSnapshot())).data)
  console.log('written', DIR)
}
