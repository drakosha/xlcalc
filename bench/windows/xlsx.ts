/**
 * The xlsx steps through Excel for Windows. Filters (`bench/coverage/xlsx/filters.ts`): each file, as the Mac step sends
 * it, computed fully and saved by Excel (`SaveAs`, xlsx) into `test/excel/win/xlsx/filters/`, with `filters.json`
 * recording what was sent (a hash per file) or Excel's refusal; the books of `bench/coverage/xlsx/stored.ts` the same
 * way into `xlsx/stored/`. Formats (`bench/coverage/xlsx/formats.ts`): the codes
 * Excel shows for the built-in number formats 1 to 58 of the Mac step's book, `NumberFormat` and `NumberFormatLocal`,
 * into `test/excel/win/xlsx/formats.json`, another locale's into `xlsx/<tag>/formats.json`. Forms
 * (`bench/coverage/xlsx/forms.ts`): the Mac step's formulas typed into its book (`Formula2`) and the book saved into
 * `xlsx/forms/`. `test/xlsx-win.test.ts` and `test/xlsx-forms.test.ts` read them.
 */
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import type { StoredRun } from '../coverage/xlsx.js'
import { colorBook } from '../coverage/xlsx/colors.js'
import { colorHeldCases, THEME_HELD } from '../coverage/xlsx/colors-held.js'
import { FILTER_FILES, filterFile } from '../coverage/xlsx/filters.js'
import { FORMAT_IDS, formatBook } from '../coverage/xlsx/formats.js'
import { FORM_CASES, FORMS_BOOK } from '../coverage/xlsx/forms.js'
import { writeXlsx } from '../../src/xlsx.js'
import { STORED_FILES, storedFile } from '../coverage/xlsx/stored.js'
import { partsHash } from '../coverage/xlsx/sent.js'
import { localeDir } from '../coverage/locales.js'
import { loadStored, secondsSince, writeAnswers } from '../coverage/stored.js'
import { winFile } from './functions.js'
import { inJobs, metaOf, OUT, runJob, type Step, type Tag, type WinEnv } from './remote.js'

/** Files of `files`, each computed fully and saved by Excel into `xlsx/<kind>/`, recorded with what was sent in `xlsx/<kind>.json`; `first`, steps before (and in the hash of what was sent). */
async function savedFiles<F extends { id: string }>(kind: string, files: readonly F[], bytes: (f: F) => Promise<Uint8Array>, tag: Tag, e: WinEnv, first: Step[] = []): Promise<void> {
  const dir = winFile(`xlsx/${kind}`)
  const file = winFile(`xlsx/${kind}.json`)
  const run: StoredRun = loadStored<StoredRun>(file) ?? { meta: {}, books: {} }
  const books = await Promise.all(
    files.map(async (c) => {
      const data = await bytes(c)
      const hash = await partsHash(data)
      return { id: c.id, data, hash: first.length ? createHash('sha256').update(hash + JSON.stringify(first)).digest('hex') : hash }
    }),
  )
  run.books = Object.fromEntries(Object.entries(run.books).filter(([id]) => books.some((b) => b.id === id)))
  const todo = books.filter((b) => !(run.books[b.id]?.sent === b.hash && (run.books[b.id].rejected || existsSync(join(dir, `${b.id}.xlsx`)))))
  mkdirSync(dir, { recursive: true })
  const t0 = Date.now()
  const save = () => {
    run.meta = { ...metaOf(e, tag), seconds: secondsSince(t0) }
    writeAnswers(file, run)
  }
  await inJobs(
    `xlsx-${kind}`,
    todo,
    async (b, f) => {
      writeFileSync(f, b.data)
      return { file: f, steps: [...first, { do: 'full' }, { do: 'saveAs' }] }
    },
    (b, a, cutShort) => {
      if (!a && cutShort) return [b]
      if (a && 'saved' in a && a.saved) {
        writeFileSync(join(dir, `${b.id}.xlsx`), a.saved)
        run.books[b.id] = { sent: b.hash }
      } else run.books[b.id] = { sent: b.hash, rejected: a ? ('rejected' in a ? a.rejected : 'not saved') : 'Excel hangs on the file' }
      return []
    },
    save,
  )
  save()
  console.log(`xlsx ${kind}: ${todo.length} files sent, ${Object.values(run.books).filter((x) => x.rejected).length} refused`)
}

export const stepXlsxFilters = (tag: Tag, e: WinEnv) => savedFiles('filters', FILTER_FILES, filterFile, tag, e)

/** The books over functions the engine computes no value for (`bench/coverage/xlsx/stored.ts`), as Excel saves them. */
export const stepXlsxStored = (tag: Tag, e: WinEnv) => savedFiles('stored', STORED_FILES, storedFile, tag, e)

/** The Mac step's typed formulas (`bench/coverage/xlsx/forms.ts`) as Excel for Windows saves them. */
export const stepXlsxForms = (tag: Tag, e: WinEnv) =>
  savedFiles('forms', [{ id: 'forms' }], async () => (await writeXlsx({ book: FORMS_BOOK })).data, tag, e, FORM_CASES.map((c) => ({ do: 'set', sheet: 'S', cell: c.range, kind: 'f2', value: c.formula })))

/** The codes Excel for Windows shows for the built-in formats 1 to 58 in `tag`: `formats` in English (`NumberFormat`), `local` in the user's language. */
export async function stepXlsxFormats(tag: Tag, e: WinEnv): Promise<void> {
  const file = winFile(`xlsx/${localeDir(tag)}formats.json`)
  const book = join(OUT, 'xlsx-formats.xlsx')
  mkdirSync(OUT, { recursive: true })
  writeFileSync(book, await formatBook())
  const t0 = Date.now()
  const [a] = await runJob('xlsx-formats', [{ file: book, steps: [{ do: 'shown', sheet: 'Formats', cells: FORMAT_IDS.map((id) => `A${id}`), codes: true }] }])
  if (!a || !('answers' in a)) throw new Error(`xlsx formats: Excel did not read the book: ${a ? a.rejected : 'the job ran past its time'}`)
  // each cell: its shown text, its colour, NumberFormat, NumberFormatLocal
  const shown = a.answers[0] as [string, number, string, string][]
  const by = (k: 2 | 3) => Object.fromEntries(FORMAT_IDS.map((id, i) => [id, shown[i][k]]))
  writeAnswers(file, { meta: { ...metaOf(e, tag), seconds: secondsSince(t0) }, formats: by(2), local: by(3) })
  console.log(`xlsx formats ${tag}: ${shown.length} read`)
}

/** The fill colours Excel for Windows shows for the book of `xlsx/colors-held.ts`, in the Mac step's shape, into `xlsx/colors-held.json`. */
export async function stepXlsxColorsHeld(tag: Tag, e: WinEnv): Promise<void> {
  const cases = colorHeldCases()
  const book = join(OUT, 'xlsx-colors-held.xlsx')
  mkdirSync(OUT, { recursive: true })
  writeFileSync(book, await colorBook(cases, THEME_HELD))
  const t0 = Date.now()
  const [a] = await runJob('xlsx-colors-held', [{ file: book, steps: [{ do: 'fills', sheet: 'Colors', cells: cases.map((c) => c.cell) }] }])
  if (!a || !('answers' in a)) throw new Error(`xlsx colors-held: Excel did not read the book: ${a ? a.rejected : 'the job ran past its time'}`)
  const hex = (bgr: number) => `#${[bgr & 255, (bgr >> 8) & 255, (bgr >> 16) & 255].map((c) => c.toString(16).padStart(2, '0')).join('').toUpperCase()}`
  const shown = a.answers[0] as [number, number][]
  writeAnswers(winFile('xlsx/colors-held.json'), { meta: { ...metaOf(e, tag), seconds: secondsSince(t0) }, colors: Object.fromEntries(cases.map((c, i) => [c.color, hex(shown[i][0])])) })
  console.log(`xlsx colors-held ${tag}: ${shown.length} read`)
}
