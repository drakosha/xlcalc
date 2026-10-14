/**
 * Row heights against Excel for Mac: the cases of `rows/<set>.ts` in a locale are written by `writeXlsx` into books of up
 * to ROWS_PART rows (`rowsBook`), a case whose column has a style alone, as that style reaches every row; Excel reads each case's row height, types its texts and fits the row
 * (`excel-rows.applescript`), reads the row's height and whether its cells wrap, and saves the book, whose `<row>`
 * elements give the height Excel stores and its flags. `npx tsx bench/coverage/run.ts rows <set> [ru|<tag>]` stores the
 * answers in `test/excel/rows/[<tag>/]<set>.json` with the day of the run; `test/rows-excel.test.ts` compares the engine
 * with them. `FROWS_ONLY=id,id` runs only those cases, `FROWS_AGAIN=id,id` those again, `FCOV_FRESH=1` every case again.
 */
import JSZip from 'jszip'
import { execFileSync } from 'node:child_process'
import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { writeXlsx } from '../../src/xlsx.js'
import { formatAddress } from '../../src/model.js'
import { declineRepair, ensureExcel, excelVersion, runWatched, settleRejected } from './excel.js'
import { localeDir, localeMeta } from './locales.js'
import { inParts } from './book.js'
import { answeredLine, chosen, inCaseOrder, keptSame, loadStored, secondsSince, setModule, took, writeAnswers } from './stored.js'
import { caseCells, caseColumn, caseRow, cellKey, rowsBook, type RowAnswer, type RowCase, type RowSet } from './rows/cases.js'
import { HERE, ROOT } from './paths.js'

const SCRIPT = join(HERE, 'excel-rows.applescript')
const KEYS = join(HERE, 'excel-keys.applescript')
/** The cases of one book, at most: a row each. */
export const ROWS_PART = 50

export interface StoredRows {
  meta: Record<string, string | number>
  cases: RowCase[]
  answers: Record<string, RowAnswer>
}

/** The stored answers of a set: `test/excel/rows/<set>.json`, in a directory of its tag for a locale other than en-US. */
export const rowsFile = (set: string, tag = 'en-US', win = false) => join(ROOT, 'test/excel', win ? 'win' : '', 'rows', localeDir(tag), `${set}.json`)

/** The stored answers of a set that its current cases keep (a case changed runs again, as do those `FROWS_AGAIN` names), with the stored file's metadata. */
export function keptRows(file: string, cases: RowCase[]): { answers: Record<string, RowAnswer>; meta?: StoredRows['meta'] } {
  const prev = loadStored<StoredRows>(file)
  const answers = keptSame({ cases }, prev)
  for (const id of process.env.FROWS_AGAIN?.split(',') ?? []) delete answers[id]
  return { answers, meta: prev?.meta }
}

/** The rows of the first sheet of a saved book with a height, as their `<row>` elements give them, by number. */
export async function savedRows(saved: Uint8Array): Promise<Map<number, NonNullable<RowAnswer['saved']>>> {
  const zip = await JSZip.loadAsync(saved)
  const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  const out = new Map<number, NonNullable<RowAnswer['saved']>>()
  for (const m of xml.matchAll(/<row\s([^>]*?)\/?>/g)) {
    const a = Object.fromEntries([...m[1].matchAll(/(\w+)="([^"]*)"/g)].map((x) => [x[1], x[2]]))
    const on = (k: string) => a[k] === '1' || a[k] === 'true'
    if (a.ht !== undefined) out.set(Number(a.r), { ht: Number(a.ht), ...(on('customHeight') ? { customHeight: true } : {}), ...(on('hidden') ? { hidden: true } : {}) })
  }
  return out
}

/** A cell of the k-th case of a part, by the case's column and the row's offset from the case's. */
export const cellOf = (k: number, col: number, dr = 0) => formatAddress(caseRow(k) - 1 + dr, caseColumn(k) + col)

/** The cases of a part as the script takes them: [row, fit, [[cell, text, via, last cell of a fill]], [cells to read]]. */
export const partSpec = (part: RowCase[]) =>
  part.map((c, k) => [
    caseRow(k),
    !c.typed || c.fitAfter ? 1 : 0,
    (c.typed ?? []).map(([col, t]) => [cellOf(k, col), t, c.via ?? 'local', cellOf(k, c.fillTo ?? col)]),
    caseCells(c).map(([dr, col]) => cellOf(k, col, dr)),
  ])

/** The answers of a part from what Excel read of each case's row and the file it saved. */
export async function partAnswers(part: RowCase[], before: number[], refused: string[][], read: [number, boolean, [string, boolean][]][], saved: Uint8Array): Promise<Record<string, RowAnswer>> {
  const rows = await savedRows(saved)
  const out: Record<string, RowAnswer> = {}
  part.forEach((c, k) => {
    const [height, standard, cells] = read[k]
    const why = Object.fromEntries(refused[k].flatMap((e, i) => (e ? [[i, e]] : [])))
    out[c.id] = {
      before: before[k],
      height,
      standard,
      saved: rows.get(caseRow(k)) ?? null,
      cells: Object.fromEntries(caseCells(c).map((x, i) => [cellKey(x), { formula: cells[i][0], wrap: cells[i][1] }])),
      ...(Object.keys(why).length ? { refused: why } : {}),
    }
  })
  return out
}

/** One part in Excel. A window that is not a book stops the run (nothing is clicked); a book Excel repairs stops it too. */
async function runPart(part: RowCase[], label: string): Promise<Record<string, RowAnswer>> {
  if (declineRepair()) console.log('closed a leftover book repair window')
  const keys = part.some((c) => c.via === 'keys' || c.via === 'fill' || c.via === 'paste')
  return runWatched<Record<string, RowAnswer>>({
    script: SCRIPT,
    label: `rows ${label}`,
    folder: 'frows',
    book: async (file) => {
      const { data, issues } = await writeXlsx({ book: rowsBook(part) }, { values: 'omit' })
      if (issues.length) throw new Error(`rows ${label}: the book does not go into a file whole: ${JSON.stringify(issues)}`)
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
      throw new Error(`rows ${label}: Excel would repair the book`)
    },
    done: (out, dir) => {
      const raw = JSON.parse(out) as { before: number[]; refused: string[][]; rows: [number, boolean, [string, boolean][]][] }
      return partAnswers(part, raw.before, raw.refused, raw.rows, new Uint8Array(readFileSync(join(dir, 'saved.xlsx'))))
    },
  })
}

/** Runs the set's cases in a locale without a stored answer for the same case (FCOV_FRESH=1 runs all) and stores the answers. */
export async function stepRows(set: string, tag: string): Promise<void> {
  const cases = (await setModule<RowSet>('rows', set)).buildRowCases(tag)
  const target = rowsFile(set, tag)
  const { answers, meta: prev } = keptRows(target, cases)
  const only = chosen('FROWS_ONLY')
  const todo = cases.filter((c) => !answers[c.id] && only(c))
  const t0 = Date.now()
  // the time of the runs before it too
  const save = () => writeAnswers(target, { meta: { version: excelVersion(), locale: localeMeta(tag), date: new Date().toISOString().slice(0, 10), seconds: secondsSince(t0 - Number(prev?.seconds ?? 0) * 1000) }, cases, answers: inCaseOrder(cases, answers) })
  // the clipboard's text, put back after a set that pastes
  const UTF8 = { ...process.env, LANG: 'en_US.UTF-8' }
  const pastes = todo.some((c) => c.via === 'paste')
  const clipboard = pastes ? execFileSync('pbpaste', { encoding: 'utf8', env: UTF8 }) : ''
  await ensureExcel(tag)
  try {
    // a column's style reaches every row of the sheet: such a case has a book of its own
    const alone = (c: RowCase) => Object.values(c.columns ?? {}).some((x) => x.look)
    const parts = [...inParts(todo.filter((c) => !alone(c)), ROWS_PART), ...todo.filter(alone).map((c) => [c])]
    for (const [k, part] of parts.entries()) {
      Object.assign(answers, await runPart(part, `${set} ${tag} part ${k + 1}`))
      save()
      console.log(`rows ${set} ${tag}: part ${k + 1}, ${part.length} cases, ${took(t0)} s`)
    }
  } finally {
    if (pastes) execFileSync('pbcopy', { input: clipboard, env: UTF8 })
    save()
    console.log(answeredLine(`rows ${set} ${tag}`, todo.length, answers, cases))
  }
}
