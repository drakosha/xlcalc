/**
 * The xlsx filters (`bench/coverage/xlsx/filters.ts`) through Excel for Windows: each file, as the Mac step sends it,
 * computed fully and saved by Excel (`SaveAs`, xlsx) into `test/excel/win/xlsx/filters/`, with `filters.json`
 * recording what was sent (a hash per file) or Excel's refusal. `test/xlsx-win.test.ts` reads what Excel saved.
 */
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import type { StoredRun } from '../coverage/xlsx.js'
import { FILTER_FILES, filterFile } from '../coverage/xlsx/filters.js'
import { winFile } from './functions.js'
import { inJobs, metaOf, type Tag, type WinEnv } from './remote.js'

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
