/**
 * Paste Special in Excel for Windows: each case of `bench/coverage/paste/<set>.ts` in the file the Mac step sends
 * (`writeXlsx`), its range copied (`Range.Copy`) and pasted at the target (`Range.PasteSpecial` with the case's
 * `XlPasteType`, operation, Skip Blanks and Transpose; a plain copy is `Range.Copy` with a destination), and the book
 * saved. Answers go to `test/excel/win/paste/<set>.json` in the Mac's shape, the saved files beside them;
 * `test/paste-win.test.ts` reads them.
 */
import { pasteSet } from '../coverage/paste.js'
import type { PasteAnswer, PasteCase, PasteOperation, PasteWhat } from '../coverage/paste/cases.js'
import { fileAnswers, secondsSince } from '../coverage/stored.js'
import { fileJobs } from './books.js'
import { winFile } from './functions.js'
import { metaOf, type Tag, type WinEnv } from './remote.js'

/** Excel's `XlPasteType` of each kind the harness pastes. */
const PASTE_TYPE: Record<Exclude<PasteWhat, 'copy'>, number> = {
  all: -4104,
  values: -4163,
  formulas: -4123,
  formats: -4122,
  validation: 6,
  allExceptBorders: 7,
  columnWidths: 8,
  formulasAndNumberFormats: 11,
  valuesAndNumberFormats: 12,
  notes: -4144,
}

/** Excel's `XlPasteSpecialOperation` of each operation. */
const OPERATION: Record<PasteOperation, number> = { add: 2, subtract: 3, multiply: 4, divide: 5 }

export async function stepPasteWin(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const { built, files, sent } = await pasteSet(set)
  const store = fileAnswers<PasteCase, PasteAnswer>(winFile(`paste/${set}.json`), built.cases, sent)
  const t0 = Date.now()
  const save = () => store.save({ ...metaOf(e, tag), seconds: secondsSince(t0) }, built)
  await fileJobs(
    `paste ${set}`,
    `paste-${set}`,
    store,
    files,
    (c) => [
      { do: 'pasteSpecial', sheet: c.from.sheet, range: c.from.range, target: c.to.sheet, to: c.to.range, ...(c.what === 'copy' ? { copy: true } : { paste: PASTE_TYPE[c.what] }), ...(c.operation ? { operation: OPERATION[c.operation] } : {}), skipBlanks: !!c.skipBlanks, transpose: !!c.transpose },
      { do: 'saveAs' },
    ],
    (_, a) => {
      const x = a.answers[0] as { error?: string } | null
      return x?.error ? { error: x.error } : {}
    },
    save,
  )
}
