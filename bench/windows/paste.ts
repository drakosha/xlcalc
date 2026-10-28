/**
 * Paste Special in Excel for Windows: each case of `bench/coverage/paste/<set>.ts` in the file the Mac step sends
 * (`writeXlsx`), its range copied (`Range.Copy`) and pasted at the target (`Range.PasteSpecial` with the case's
 * `XlPasteType`, operation, Skip Blanks and Transpose; a plain copy is `Range.Copy` with a destination), and the book
 * saved. The Format Painter's cases need a session on the screen (`screen.ts`): `painter.ps1` runs them there, the
 * source selected, the painter turned on and the target selected with the keyboard as on the Mac; without a session on
 * the screen they wait. Answers go to `test/excel/win/paste/<set>.json` in the Mac's shape, the saved files beside
 * them; `test/paste-win.test.ts` reads them.
 */
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { pasteSet } from '../coverage/paste.js'
import { painterSelection, type PasteAnswer, type PasteCase, type PasteOperation, type PasteWhat } from '../coverage/paste/cases.js'
import { fileAnswers, secondsSince, type FileAnswers } from '../coverage/stored.js'
import { fileJobs } from './books.js'
import { winFile } from './functions.js'
import { HOST, metaOf, type Tag, type WinEnv } from './remote.js'
import { onScreen, runOnScreen, scp } from './screen.js'

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
  const dir = mkdtempSync(join(tmpdir(), 'xlwin-painter-'))
  for (const c of cases) writeFileSync(join(dir, `${c.id}.xlsx`), files.get(c.id)!)
  const job = cases.map((c) => {
    const { select, key } = painterSelection(c.to.range)
    return { file: `${c.id}.xlsx`, out: `${c.id}-saved.xlsx`, from: [c.from.sheet, c.from.range], to: [c.to.sheet, select], key }
  })
  writeFileSync(join(dir, 'job.json'), JSON.stringify(job))
  await runOnScreen('painter', 'painter.ps1', dir, 60 + 15 * cases.length)
  scp(`${HOST}:C:/xlcalc/painter/out.json`, join(dir, 'out.json'))
  const raw = JSON.parse(readFileSync(join(dir, 'out.json'), 'utf8').replace(/^\uFEFF/, '')) as { error?: string; answers: { error: string }[] }
  if (raw.error) throw new Error(`painter: ${raw.error}`)
  for (const [i, c] of cases.entries()) {
    const err = raw.answers[i]?.error
    // a painter that did not apply is the step's failure, not Excel's answer
    if (err === undefined || err) {
      console.log(`${c.id}: ${err ?? 'no answer'}`)
      continue
    }
    scp(`${HOST}:C:/xlcalc/painter/${c.id}-saved.xlsx`, join(dir, `${c.id}-saved.xlsx`))
    store.got(c.id, { answer: {}, saved: new Uint8Array(readFileSync(join(dir, `${c.id}-saved.xlsx`))) })
  }
}

export async function stepPasteWin(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const { built, files, sent } = await pasteSet(set)
  const store = fileAnswers<PasteCase, PasteAnswer>(winFile(`paste/${set}.json`), built.cases, sent)
  const t0 = Date.now()
  const save = () => store.save({ ...metaOf(e, tag), seconds: secondsSince(t0) }, built)
  const painted = store.todo.filter((c) => c.what === 'painter')
  await fileJobs(
    `paste ${set}`,
    `paste-${set}`,
    { ...store, todo: store.todo.filter((c) => c.what !== 'painter') },
    files,
    (c) => [
      { do: 'pasteSpecial', sheet: c.from.sheet, range: c.from.range, target: c.to.sheet, to: c.to.range, ...(c.what === 'copy' ? { copy: true } : { paste: PASTE_TYPE[c.what as keyof typeof PASTE_TYPE] }), ...(c.operation ? { operation: OPERATION[c.operation] } : {}), skipBlanks: !!c.skipBlanks, transpose: !!c.transpose },
      { do: 'saveAs' },
    ],
    (_, a) => {
      const x = a.answers[0] as { error?: string } | null
      return x?.error ? { error: x.error } : {}
    },
    save,
  )
  if (!painted.length) return
  if (!onScreen()) return void console.log(`paste ${set}: ${painted.length} cases of the Format Painter wait for a session on the screen`)
  await paintOnScreen(painted, files, store)
  save()
  console.log(`paste ${set}: the Format Painter, ${painted.filter((c) => store.answers[c.id]).length} of ${painted.length} answered`)
}
