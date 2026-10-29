/**
 * The order of the keys of a session's workbook (a sheet's fields, its cells) is not part of it: a fill that gives rows
 * a look puts `looks` among a sheet's fields in another place in a thread than in a synchronous session, and undo and
 * redo put them back in other places again. The workbook by content, its hash, the file written from it and the
 * changes are the same in every kind of session and after undo and redo.
 */
import { createHash } from 'node:crypto'
import { afterAll, describe, expect, it } from 'vitest'
import type { AsyncSession, BookInput, Clock, StructureOp } from '../src/index.js'
import { writeXlsx } from '../src/xlsx.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const clock: Clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

/** A fill of numbers over rows of a style on a sheet with validation and looks; a fill of formulas around cells out of order. */
const CASES: [string, BookInput, StructureOp][] = [
  [
    'a fill over styled rows',
    { styles: [{}, { font: { bold: true } }, { font: { italic: true } }], sheets: [{ name: 'S', cells: { D4: { value: 1 }, D5: { value: 2 } }, rows: { '9:12': { style: 1 } }, validations: [{ range: 'A1', type: 'list', formula1: '"a,b"' }], looks: [{ range: 'G7', style: 2 }] }] },
    { op: 'fillRange', sheet: 'S', range: 'D4:D5', to: 'D4:D12', type: 'series' },
  ],
  ['a fill of formulas', { sheets: [{ name: 'S', cells: { C3: { value: 3 }, A1: { value: 1 }, B2: { formula: 'A1*2' }, A9: { value: 9 } } }] }, { op: 'fillRange', sheet: 'S', range: 'A1:B2', to: 'A1:B8' }],
]

const sha = (d: Uint8Array) => createHash('sha256').update(d).digest('hex')

/** What does not depend on the order of keys: the workbook by content, its hash, the file written from it, the cells of the changes in order. */
async function state(s: AsyncSession, changes?: { cell: string }[]) {
  const book = await s.getBook()
  return { book, hash: await s.getHash(), file: sha((await writeXlsx({ book })).data), changes: changes?.map((c) => c.cell) }
}

describe.each(CASES)('%s', (_, book, op) => {
  it('gives the same workbook, hash, file and changes in every kind of session, after the operation, undo and redo', async () => {
    const all: Awaited<ReturnType<typeof state>>[][] = []
    for (const kind of KINDS) {
      const s = await (await opener(kind))(structuredClone(book), { clock })
      const r = await s.editStructure(op)
      if (!r.ok) throw new Error(r.message)
      const steps = [await state(s, r.changes)]
      steps.push(await state(s, (await s.undo())!.changes), await state(s, (await s.redo())!.changes))
      all.push(steps)
    }
    for (const steps of all.slice(1)) expect(steps).toEqual(all[0])
    expect(all[0][2]).toEqual(all[0][0])
  })
})
