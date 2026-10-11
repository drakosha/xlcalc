/**
 * Formats of rows and columns in Excel for Windows: each case of `bench/coverage/lines/<set>.ts` in the file the Mac step
 * sends (`writeXlsx`), its steps through COM (`NumberFormat`, `Font.Bold`, `Font.Italic`, `Interior.Color` of a range, `Value2` of a
 * cell, `ClearContents`, `ClearFormats` and `Clear` of a range), the book computed, the read cells' `Range.Text`, number format and display format (`DisplayFormat`: fill and font
 * colours, bold, italic), and the book saved. Answers go to `test/excel/win/lines/<set>.json` in the Mac's shape, the
 * saved files beside them; `test/lines-excel.test.ts` reads them.
 */
import { fromBgr } from '../coverage/cond/cases.js'
import { linesSet } from '../coverage/lines.js'
import type { LineLook, LinesAnswer, LinesCase, LineStep } from '../coverage/lines/cases.js'
import { fileAnswers, secondsSince } from '../coverage/stored.js'
import { comEdit, fileJobs, stepErrors } from './books.js'
import { winFile } from './functions.js'
import { metaOf, type Step, type Tag, type WinEnv } from './remote.js'

type RawLook = [address: string, text: string, fill: number, colorIndex: number, font: number, bold: boolean, italic: boolean, format: string]

/** xlNone: no fill. */
const NO_FILL = -4142

/** `#RRGGBB` as COM's BGR. */
const bgr = (c: string) => parseInt(c.slice(5, 7) + c.slice(3, 5) + c.slice(1, 3), 16)

/** A step through COM: the range's formats, a value typed into a cell, or the range cleared. */
const comStep = (s: LineStep): Step =>
  s.clear
    ? { do: 'clear', sheet: s.sheet, range: s.range, what: s.clear }
    : s.value !== undefined
    ? { do: 'set', sheet: s.sheet, cell: s.range, ...comEdit(s.value) }
    : { do: 'format', sheet: s.sheet, range: s.range, ...(s.format !== undefined ? { numberFormat: s.format } : {}), ...(s.bold ? { bold: true } : {}), ...(s.italic ? { italic: true } : {}), ...(s.fill ? { fill: bgr(s.fill) } : {}) }

export async function stepLines(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const { built, files, sent } = await linesSet(set)
  const store = fileAnswers<LinesCase, LinesAnswer>(winFile(`lines/${set}.json`), built.cases, sent)
  const t0 = Date.now()
  const save = () => store.save({ ...metaOf(e, tag), seconds: secondsSince(t0) }, built)
  await fileJobs(
    `lines ${set}`,
    `lines-${set}`,
    store,
    files,
    (c) => [...(c.steps ?? []).map(comStep), { do: 'full' }, { do: 'cond', kind: 'looks', read: c.read, codes: true }, { do: 'saveAs' }],
    (c, a) => {
      const n = (c.steps ?? []).length
      const errors = stepErrors(a.answers.slice(0, n), (i) => c.steps![i].range)
      const looks: LinesAnswer['looks'] = {}
      for (const [sheet, cells] of a.answers[n + 1] as [string, RawLook[]][])
        looks[sheet] = Object.fromEntries(cells.map(([addr, text, fill, ci, font, bold, italic, format]): [string, LineLook] => [addr, { text, format, fill: ci === NO_FILL ? null : fromBgr(fill), font: fromBgr(font), bold, italic }]))
      return { ...(errors.length ? { error: errors.join('; ') } : {}), looks }
    },
    save,
  )
}
