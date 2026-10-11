/**
 * Column widths against Excel for Mac: the cases of `fit/<set>.ts` in a locale are written by `writeXlsx` into books of up
 * to PART columns (`fitBook`); Excel fits each case's column or types its texts (`excel-fit.applescript`), reads the
 * column's width and what its cells show, and saves the book, whose `<col>` elements give the width Excel stores and its
 * flags. `npx tsx bench/coverage/run.ts fit <set> [ru|<tag>]` stores the answers in `test/excel/fit/[<tag>/]<set>.json`
 * with the day of the run; `test/fit-excel.test.ts` compares the engine with them. `FFIT_ONLY=id,id` runs only those
 * cases, `FFIT_AGAIN=id,id` those again, `FCOV_FRESH=1` every case again.
 */
import JSZip from 'jszip'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { writeXlsx } from '../../src/xlsx.js'
import { indexToCol } from '../../src/model.js'
import { declineRepair, ensureExcel, excelVersion, runWatched, settleRejected } from './excel.js'
import { localeDir, localeMeta } from './locales.js'
import { inCaseOrder } from './stored.js'
import { caseColumn, caseRows, fitBook, type FitAnswer, type FitCase, type FitSet } from './fit/cases.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
const SCRIPT = join(HERE, 'excel-fit.applescript')
const KEYS = join(HERE, 'excel-keys.applescript')
const PART = 60

export interface StoredFit {
  meta: Record<string, string | number>
  cases: FitCase[]
  answers: Record<string, FitAnswer>
}

/** The stored answers of a set: `test/excel/fit/<set>.json`, in a directory of its tag for a locale other than en-US. */
export const fitFile = (set: string, tag = 'en-US', win = false) => join(ROOT, 'test/excel', win ? 'win' : '', 'fit', localeDir(tag), `${set}.json`)

/**
 * The stored answers of a set that its current cases keep (a case changed runs again, as do the cases `FFIT_AGAIN=id,id`
 * names, whose book the writer writes otherwise now; FCOV_FRESH=1 runs all), with the stored file's metadata.
 */
export function keptAnswers(file: string, cases: FitCase[]): { answers: Record<string, FitAnswer>; meta?: StoredFit['meta'] } {
  const prev = existsSync(file) && !process.env.FCOV_FRESH ? (JSON.parse(readFileSync(file, 'utf8')) as StoredFit) : null
  const was = new Map(prev?.cases.map((c) => [c.id, JSON.stringify(c)]) ?? [])
  const again = new Set(process.env.FFIT_AGAIN?.split(','))
  const answers: Record<string, FitAnswer> = {}
  for (const c of cases) if (prev?.answers[c.id] && was.get(c.id) === JSON.stringify(c) && !again.has(c.id)) answers[c.id] = prev.answers[c.id]
  return { answers, meta: prev?.meta }
}

/** The columns of the first sheet of a saved book as its `<col>` elements give them, by letter. */
export async function savedColumns(saved: Uint8Array): Promise<Map<string, FitAnswer['saved']>> {
  const zip = await JSZip.loadAsync(saved)
  const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  const out = new Map<string, FitAnswer['saved']>()
  for (const m of xml.matchAll(/<col\s([^>]*?)\/?>/g)) {
    const a = Object.fromEntries([...m[1].matchAll(/(\w+)="([^"]*)"/g)].map((x) => [x[1], x[2]]))
    const on = (k: string) => a[k] === '1' || a[k] === 'true'
    for (let c = Number(a.min); c <= Number(a.max); c++)
      out.set(indexToCol(c - 1), { width: Number(a.width), ...(on('customWidth') ? { customWidth: true } : {}), ...(on('bestFit') ? { bestFit: true } : {}), ...(on('hidden') ? { hidden: true } : {}) })
  }
  return out
}

/** The cases of a part as the script takes them: [column, fit first, [[row, text, via]], [rows to read]]. */
export const partSpec = (part: FitCase[]) => part.map((c, k) => [indexToCol(caseColumn(k)), c.fitFirst ? 1 : 0, (c.typed ?? []).map(([r, t]) => [r, t, c.via ?? 'local']), caseRows(c)])

/** The answers of a part from what Excel read of each case's column and the file it saved. */
export async function partAnswers(part: FitCase[], refused: string[][], read: [number, number, boolean, string[]][], saved: Uint8Array): Promise<Record<string, FitAnswer>> {
  const cols = await savedColumns(saved)
  const out: Record<string, FitAnswer> = {}
  part.forEach((c, k) => {
    const [columnWidth, points, standard, shown] = read[k]
    const rows = caseRows(c)
    const why = Object.fromEntries(refused[k].flatMap((e, i) => (e ? [[i, e]] : [])))
    out[c.id] = {
      columnWidth,
      points,
      standard,
      saved: cols.get(indexToCol(caseColumn(k))) ?? null,
      texts: Object.fromEntries(rows.map((r, i) => [r, shown[i]])),
      ...(Object.keys(why).length ? { refused: why } : {}),
    }
  })
  return out
}

/** One part in Excel. A window that is not a book stops the run (nothing is clicked); a book Excel repairs stops it too. */
async function runPart(part: FitCase[], label: string): Promise<Record<string, FitAnswer>> {
  if (declineRepair()) console.log('closed a leftover book repair window')
  const keys = part.some((c) => c.via === 'keys')
  return runWatched<Record<string, FitAnswer>>({
    script: SCRIPT,
    label: `fit ${label}`,
    folder: 'ffit',
    book: async (file) => {
      const { data, issues } = await writeXlsx({ book: fitBook(part) }, { values: 'omit' })
      if (issues.length) throw new Error(`fit ${label}: the book does not go into a file whole: ${JSON.stringify(issues)}`)
      writeFileSync(file, data)
    },
    args: (dir) => [join(dir, 'saved.xlsx'), 'saved.xlsx', JSON.stringify(partSpec(part)), KEYS],
    timeout: 290_000,
    poll: 1000,
    after: 'part',
    // typed on the keyboard, the script answers the window a refused formula brings itself
    fatal: !keys,
    repair: async (stop) => {
      await stop()
      await settleRejected()
      throw new Error(`fit ${label}: Excel would repair the book`)
    },
    done: (out, dir) => {
      const raw = JSON.parse(out) as { refused: string[][]; cases: [number, number, boolean, string[]][] }
      return partAnswers(part, raw.refused, raw.cases, new Uint8Array(readFileSync(join(dir, 'saved.xlsx'))))
    },
  })
}

/** Runs the set's cases in a locale without a stored answer for the same case (FCOV_FRESH=1 runs all) and stores the answers. */
export async function stepFit(set: string, tag: string): Promise<void> {
  if (!/^[a-z0-9-]+$/.test(set)) throw new Error(`set name: [a-z0-9-], not ${set}`)
  const mod = (await import(`./fit/${set}.js`)) as FitSet
  const cases = mod.buildFitCases(tag)
  const target = fitFile(set, tag)
  const { answers, meta: prev } = keptAnswers(target, cases)
  const only = process.env.FFIT_ONLY?.split(',')
  const todo = cases.filter((c) => !answers[c.id] && (!only || only.includes(c.id)))
  const t0 = Date.now()
  mkdirSync(join(target, '..'), { recursive: true })
  const save = () => {
    const meta = { version: excelVersion(), locale: localeMeta(tag), date: new Date().toISOString().slice(0, 10), seconds: Math.round((Number(prev?.seconds ?? 0) * 1000 + Date.now() - t0) / 100) / 10 }
    writeFileSync(target, JSON.stringify({ meta, cases, answers: inCaseOrder(cases, answers) }, null, 1) + '\n')
  }
  await ensureExcel(tag)
  try {
    for (let k = 0; k * PART < todo.length; k++) {
      const part = todo.slice(k * PART, (k + 1) * PART)
      Object.assign(answers, await runPart(part, `${set} ${tag} part ${k + 1}`))
      save()
      console.log(`fit ${set} ${tag}: part ${k + 1}, ${part.length} cases, ${((Date.now() - t0) / 1000).toFixed(1)} s`)
    }
  } finally {
    save()
    console.log(`fit ${set} ${tag}: ${todo.length} cases run, ${Object.keys(answers).length} of ${cases.length} answered`)
  }
}
