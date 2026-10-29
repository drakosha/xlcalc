/**
 * Paste Special in Excel for Windows: each case of `bench/coverage/paste/<set>.ts` in the file the Mac step sends
 * (`writeXlsx`), its range copied (`Range.Copy`) and pasted at the target (`Range.PasteSpecial` with the case's
 * `XlPasteType`, operation, Skip Blanks and Transpose; a plain copy is `Range.Copy` with a destination), and the book
 * saved; a case that copies from another workbook sends that book with the step, opened beside. The Format Painter's cases need a session on the screen (`screen.ts`): `painter.ps1` runs them there, the
 * source selected, the painter turned on and the target selected with the keyboard as on the Mac; without a session on
 * the screen they wait. The formula of each cell a case names (`local`) is read in the user's language
 * (`FormulaLocal`). Answers go to `test/excel/win/paste/[<tag>/]<set>.json` in the Mac's shape, the saved files beside
 * them; `test/paste-win.test.ts` reads them.
 */
import { localeDir } from '../coverage/locales.js'
import { pasteAnswer, pasteSet } from '../coverage/paste.js'
import { painterSelection, type PasteAnswer, type PasteCase, type PasteOperation, type PasteWhat } from '../coverage/paste/cases.js'
import { fileAnswers, secondsSince, type FileAnswers } from '../coverage/stored.js'
import { fileJobs } from './books.js'
import { winFile } from './functions.js'
import { metaOf, type Tag, type WinEnv } from './remote.js'
import { onScreen, savedOnScreen } from './screen.js'

/** Excel's `XlPasteType` of each kind the harness pastes. */
const PASTE_TYPE: Record<Exclude<PasteWhat, 'copy' | 'painter'>, number> = {
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
  allUsingSourceTheme: 13,
  allMergingConditionalFormats: 14,
}

/** Excel's `XlPasteSpecialOperation` of each operation. */
const OPERATION: Record<PasteOperation, number> = { add: 2, subtract: 3, multiply: 4, divide: 5 }

/** The Format Painter's cases of the store still to do, on the screen; each a book saved, or left to do where the painter did not apply. */
async function paintOnScreen(cases: PasteCase[], files: Map<string, Uint8Array>, store: FileAnswers<PasteCase, PasteAnswer>): Promise<void> {
  const byId = new Map(cases.map((c) => [c.id, c]))
  const saved = await savedOnScreen('painter', 'painter.ps1', new Map(cases.map((c) => [c.id, files.get(c.id)!])), (id) => {
    const c = byId.get(id)!
    const { select, key } = painterSelection(c.to.range)
    const more = (c.more ?? []).map((t) => {
      const q = painterSelection(t.range)
      return { to: [t.sheet, q.select], key: q.key }
    })
    return { from: [c.from.sheet, c.from.range], to: [c.to.sheet, select], key, more }
  }, 60 + 15 * cases.length)
  for (const [id, { saved: data }] of saved) store.got(id, { answer: {}, saved: data })
}

export async function stepPasteWin(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const { built, files, sources, sent } = await pasteSet(set)
  const store = fileAnswers<PasteCase, PasteAnswer>(winFile(`paste/${localeDir(tag)}${set}.json`), built.cases, sent)
  const t0 = Date.now()
  const save = () => store.save({ ...metaOf(e, tag), seconds: secondsSince(t0) }, built)
  const painted = store.todo.filter((c) => c.what === 'painter')
  await fileJobs(
    `paste ${set}`,
    `paste-${set}`,
    { ...store, todo: store.todo.filter((c) => c.what !== 'painter') },
    files,
    (c) => [
      { do: 'pasteSpecial', sheet: c.from.sheet, range: c.from.range, target: c.to.sheet, to: c.to.range, ...(c.what === 'copy' ? { copy: true } : { paste: PASTE_TYPE[c.what as keyof typeof PASTE_TYPE] }), ...(c.operation ? { operation: OPERATION[c.operation] } : {}), skipBlanks: !!c.skipBlanks, transpose: !!c.transpose, ...(sources.has(c.id) ? { source: Buffer.from(sources.get(c.id)!).toString('base64') } : {}) },
      ...(c.local ?? []).map((cell) => ({ do: 'formulasLocal', sheet: c.to.sheet, range: cell })),
      { do: 'saveAs' },
    ],
    (c, a) => pasteAnswer(c, (a.answers[0] as { error?: string } | null)?.error, a.answers.slice(1, 1 + (c.local?.length ?? 0)).map((x) => String((x as unknown[][])[0][0]))),
    save,
  )
  if (!painted.length) return
  if (!onScreen()) return void console.log(`paste ${set}: ${painted.length} cases of the Format Painter wait for a session on the screen`)
  await paintOnScreen(painted, files, store)
  save()
  console.log(`paste ${set}: the Format Painter, ${painted.filter((c) => store.answers[c.id]).length} of ${painted.length} answered`)
}
