/**
 * Hidden sheets in Excel for Windows: each case of `bench/coverage/hidden/<set>.ts` in the file the Mac step sends, the
 * visibility it sets through COM (`Worksheet.Visible`), its sheet operation, each sheet's visibility, the active sheet
 * and the read cells' values (`Value2`) read, and the book saved. Answers go to `test/excel/win/hidden/<set>.json` in the
 * Mac's shape, the saved files beside them; `test/hidden-sheets-win.test.ts` reads them.
 */
import type { CellValue } from '../../src/contract.js'
import { hiddenSet } from '../coverage/hidden.js'
import type { HiddenAnswer, HiddenCase, Visible } from '../coverage/hidden/cases.js'
import { fileAnswers, secondsSince } from '../coverage/stored.js'
import { fileJobs, opSteps, stepErrors } from './books.js'
import { winFile } from './functions.js'
import { metaOf, type Step, type Tag, type WinEnv } from './remote.js'

/** COM's error codes of the values read (`0x800A0000` plus Excel's error number): none is expected, the read formulas catch them. */
const value = (v: unknown): CellValue => (v !== null && typeof v === 'object' ? { error: 'VALUE' } : (v as CellValue))

export async function stepHidden(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const { built, files, sent } = await hiddenSet(set)
  const store = fileAnswers<HiddenCase, HiddenAnswer>(winFile(`hidden/${set}.json`), built.cases, sent)
  const t0 = Date.now()
  const save = () => store.save({ ...metaOf(e, tag), seconds: secondsSince(t0) }, built)
  const { sheet: readSheet, cells } = built.read
  await fileJobs(
    `hidden ${set}`,
    `hidden-${set}`,
    store,
    files,
    (c) => [{ do: 'book' }, ...(c.make ?? []).map((m): Step => ({ do: 'visible', sheet: m.sheet, visible: m.visible })), ...opSteps(c), { do: 'calc' }, { do: 'book' }, ...cells.map((cell) => ({ do: 'values', sheet: readSheet, range: cell })), { do: 'saveAs' }],
    (c, a) => {
      const makes = (c.make ?? []).length
      const n = 1 + makes + (c.op ? 1 : 0)
      // the first step Excel refused, where the Mac stops
      const [error] = stepErrors(a.answers.slice(1, n), (i) => (i < makes ? 'make' : 'op'))
      const [opened] = a.answers[0] as [string, [string, number][]]
      const [active, sheets] = a.answers[n + 1] as [string, [string, number][]]
      const values = cells.map((_, i) => value((a.answers[n + 2 + i] as unknown[][])[0][0]))
      return { ...(error ? { error } : {}), opened, active, sheets: sheets.map(([name, v]) => [name, v as Visible]), values }
    },
    save,
  )
}
