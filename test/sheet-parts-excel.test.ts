/**
 * Notes, hyperlinks and frozen panes against Excel 16 for Mac (`test/excel/parts/`, see `test/sheet-parts-excel.ts`):
 * Excel read in each book what the engine holds after the case (the book as written, with what Excel made, after the
 * structure operation), and the reader reads each file Excel saved as Excel read it. Known differences are pinned in
 * `PARTS_KNOWN` and `PARTS_READ_KNOWN` (`test/sheet-parts-excel.ts`) with their lines and reasons.
 */
import { describe, expect, it } from 'vitest'
import { checkSets, PARTS_KNOWN, PARTS_READ_KNOWN, partsSets, THREAD_ISSUES } from './sheet-parts-excel.js'

const SETS = partsSets('parts/')

describe('notes, hyperlinks and frozen panes against Excel for Mac', () => {
  it('every set is answered', () => expect(SETS.map((s) => s.name)).toEqual(['files', 'frozen-held', 'struct-held', 'threads']))
  it('every pinned difference names a case', () => {
    const keys = new Set(SETS.flatMap((s) => s.cases.map((c) => `${s.name}:${c.id}`)))
    expect([...Object.keys(PARTS_KNOWN), ...Object.keys(PARTS_READ_KNOWN)].filter((k) => !keys.has(k))).toEqual([])
  })
})

checkSets(SETS, PARTS_KNOWN, PARTS_READ_KNOWN, (name) => import(`../bench/coverage/parts/${name}.ts`), THREAD_ISSUES)
