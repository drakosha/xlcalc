/**
 * The xlsx steps through Excel for Windows. Filters (`bench/coverage/xlsx/filters.ts`): each file, as the Mac step sends
 * it, computed fully and saved by Excel (`SaveAs`, xlsx) into `test/excel/win/xlsx/filters/`, with `filters.json`
 * recording what was sent (a hash per file) or Excel's refusal. Formats (`bench/coverage/xlsx/formats.ts`): the codes
 * Excel shows for the built-in number formats 1 to 58 of the Mac step's book, `NumberFormat` and `NumberFormatLocal`,
 * into `test/excel/win/xlsx/formats.json`, another locale's into `xlsx/<tag>/formats.json`. `test/xlsx-win.test.ts`
 * reads them.
 */
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import type { StoredRun } from '../coverage/xlsx.js'
import { FILTER_FILES, filterFile } from '../coverage/xlsx/filters.js'
import { FORMAT_IDS, formatBook } from '../coverage/xlsx/formats.js'
import { localeDir } from '../coverage/locales.js'
import { winFile } from './functions.js'
import { inJobs, metaOf, OUT, runJob, type Tag, type WinEnv } from './remote.js'

const sha = (x: Uint8Array) => createHash('sha256').update(x).digest('hex')

export async function stepXlsxFilters(tag: Tag, e: WinEnv): Promise<void> {
  const dir = winFile('xlsx/filters')
  const file = winFile('xlsx/filters.json')
  const run: StoredRun = existsSync(file) && !process.env.FCOV_FRESH ? (JSON.parse(readFileSync(file, 'utf8')) as StoredRun) : { meta: {}, books: {} }
  const books = await Promise.all(FILTER_FILES.map(async (c) => ({ id: c.id, data: await filterFile(c) })))
  run.books = Object.fromEntries(Object.entries(run.books).filter(([id]) => books.some((b) => b.id === id)))
  const todo = books.filter((b) => !(run.books[b.id]?.sent === sha(b.data) && (run.books[b.id].rejected || existsSync(join(dir, `${b.id}.xlsx`)))))
  mkdirSync(dir, { recursive: true })
  const t0 = Date.now()
  const save = () => {
    run.meta = { ...metaOf(e, tag), seconds: Math.round((Date.now() - t0) / 100) / 10 }
    writeFileSync(file, JSON.stringify(run, null, 1) + '\n')
  }
  await inJobs(
    'xlsx-filters',
    todo,
    async (b, f) => {
      writeFileSync(f, b.data)
      return { file: f, steps: [{ do: 'full' }, { do: 'saveAs' }] }
    },
    (b, a, cutShort) => {
      if (!a && cutShort) return [b]
      if (a && 'saved' in a && a.saved) {
        writeFileSync(join(dir, `${b.id}.xlsx`), a.saved)
        run.books[b.id] = { sent: sha(b.data) }
      } else run.books[b.id] = { sent: sha(b.data), rejected: a ? ('rejected' in a ? a.rejected : 'not saved') : 'Excel hangs on the file' }
      return []
    },
    save,
  )
  save()
  console.log(`xlsx filters: ${todo.length} files sent, ${Object.values(run.books).filter((x) => x.rejected).length} refused`)
}

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
  mkdirSync(dirname(file), { recursive: true })
  writeFileSync(file, JSON.stringify({ meta: { ...metaOf(e, tag), seconds: Math.round((Date.now() - t0) / 100) / 10 }, formats: by(2), local: by(3) }, null, 1) + '\n')
  console.log(`xlsx formats ${tag}: ${shown.length} read`)
}
