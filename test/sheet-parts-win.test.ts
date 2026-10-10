/**
 * Notes, hyperlinks and frozen panes against Excel 2024 for Windows (`test/excel/win/parts/`, `bench/windows/parts.ts`),
 * as `sheet-parts-excel.test.ts` checks the Mac's: Excel read what the engine holds after each case, and the reader reads
 * each file Excel saved as Excel read it, with the Mac's pinned differences and one of Windows's own.
 */
import { describe, expect, it } from 'vitest'
import { checkSets, PARTS_KNOWN, PARTS_READ_KNOWN, partsSets, THREAD_ISSUES } from './sheet-parts-excel.js'

const SETS = partsSets('win/parts/')

/**
 * Excel for Windows saves a relative address relative to where it saves the file, and the harness saves into a folder
 * below the book's (`out\`): `other.xlsx` comes back as `../other.xlsx`, the same file.
 */
const relative = (cell: string) => `S link ${cell}: Excel {"url":"other.xlsx"}, book {"url":"../other.xlsx"}`

const READ_KNOWN: Readonly<Record<string, string[]>> = Object.fromEntries(
  Object.entries(PARTS_READ_KNOWN).map(([k, lines]) => [k, [...lines, relative(k === 'files:after-insert' ? 'D13' : 'D12')].sort()]),
)

describe('notes, hyperlinks and frozen panes against Excel for Windows', () => {
  it('every set is answered', () => expect(SETS.map((s) => s.name)).toEqual(['files', 'frozen-held', 'struct-held', 'threads']))
})

checkSets(SETS, PARTS_KNOWN, READ_KNOWN, (name) => import(`../bench/coverage/parts/${name}.ts`), THREAD_ISSUES)
