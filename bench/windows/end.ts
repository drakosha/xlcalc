/**
 * Excel's `Range.End` in Excel for Windows (`bench/coverage/end/`): each book of a set written by `writeXlsx`, opened,
 * computed, its autofilter applied (`Range.AutoFilter`) and its texts typed (`Formula2Local`); then each case's cell's
 * `End` in its direction (`xlUp`, `xlDown`, `xlToLeft`, `xlToRight`), the address of the cell it comes to. The cases of
 * one book run in one file. The answers go to `test/excel/win/end/[<tag>/]<set>.json`; a case Excel refuses keeps its error.
 */
import type { EndDirection } from '../../src/contract.js'
import { endFile, sentBook, type EndSet } from '../coverage/end/cases.js'
import { setModule } from '../coverage/stored.js'
import type { Tag, WinEnv } from './remote.js'
import { stepByBook } from './by-book.js'

/** Excel's `XlDirection`. */
const XL_DIRECTION: Record<EndDirection, number> = { up: -4162, down: -4121, left: -4159, right: -4161 }

export async function stepEnd(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const { BOOKS, buildEndCases } = await setModule<EndSet>('end', set)
  await stepByBook({ kind: 'end', set, tag, e, books: BOOKS, cases: buildEndCases(), file: endFile(set, tag), sent: sentBook, step: (c) => ({ do: 'end', sheet: 'S', cell: c.cell, direction: XL_DIRECTION[c.direction] }), answer: (x) => (x && typeof x === 'object' && 'error' in x ? { error: String(x.error) } : { cell: String(x) }) })
}
