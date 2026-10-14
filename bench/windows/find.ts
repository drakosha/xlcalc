/**
 * Excel's Find in Excel for Windows (`bench/coverage/find/`): each book of a set written by `writeXlsx`, opened, computed,
 * its autofilter applied (`Range.AutoFilter`) and its texts typed (`Formula2Local`); then each case's query through
 * `Range.Find` from after the range's last cell and `FindNext` until Find comes back to a cell found before, at most
 * `MAX_FOUND` cells, each with its `Formula`, `FormulaLocal` and `Text`. The cases of one book run in one file. The
 * answers go to `test/excel/win/find/[<tag>/]<set>.json`; a case Find refuses keeps its error.
 */
import { MAX_FOUND, findFile, sentBook, type FindAnswer, type FindSet } from '../coverage/find/cases.js'
import { setModule } from '../coverage/stored.js'
import type { Tag, WinEnv } from './remote.js'
import { stepByBook } from './by-book.js'

export async function stepFind(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const { BOOKS, buildFindCases } = await setModule<FindSet>('find', set)
  await stepByBook({ kind: 'find', set, tag, e, books: BOOKS, cases: buildFindCases(), file: findFile(set, tag), sent: sentBook, step: (c) => ({ do: 'find', sheet: 'S', max: MAX_FOUND, ...c.query }), answer: answerOf })
}

/** A query's answer as `xl.ps1` gives it: `[[address, Formula, FormulaLocal, Text]...], more, FindNext's error, the cell it came back to`, or Find's error. */
function answerOf(x: unknown): FindAnswer {
  if (x && typeof x === 'object' && 'error' in x) return { error: String(x.error) }
  const [found, more, next, back] = x as [[string, string, string, string][], boolean, string, string?]
  return { found: found.map(([cell, formula, formulaLocal, text]) => ({ cell, formula, formulaLocal, text })), ...(more ? { more: true as const } : {}), ...(next ? { next } : {}), ...(back ? { back } : {}) }
}
