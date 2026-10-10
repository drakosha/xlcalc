/**
 * Formats of rows and columns against Excel for Mac en-US: `npx tsx bench/coverage/run.ts lines <set>` writes each book
 * of `lines/<set>.ts` with `writeXlsx`, Excel opens it, takes the case's steps, computes the book, reads what the listed
 * cells show and their number format, and saves the book (`excel-lines.applescript`). The answers go to
 * `test/excel/lines/<set>.json` and the saved files to `test/excel/lines/<set>/<id>.xlsx`; `test/lines-excel.test.ts`
 * compares them with the engine and the reader without Excel. `FLINES_ONLY=id,id` runs only those cases,
 * `FCOV_FRESH=1` every case again.
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { writeXlsx } from '../../src/xlsx.js'
import { fromRgb } from './cond/cases.js'
import { ensureExcel, excelVersion, runSavingScript } from './excel.js'
import type { LineLook, LinesAnswer, LinesCase, LinesCases, LinesSet } from './lines/cases.js'
import { inCaseOrder, keptAnswers } from './stored.js'
import { partsHash } from './xlsx/sent.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
const SCRIPT = join(HERE, 'excel-lines.applescript')

export interface StoredLines extends LinesCases {
  meta: Record<string, string | number>
  /** What was sent for each case: `partsHash` of the file `writeXlsx` wrote. */
  sent: Record<string, string>
  answers: Record<string, LinesAnswer>
  /** Cases whose book Excel did not load. */
  rejected?: Record<string, string>
}

/** The set's cases, the file of each case and what identifies it. */
export async function linesSet(set: string): Promise<{ built: LinesCases; files: Map<string, Uint8Array>; sent: Record<string, string> }> {
  if (!/^[a-z0-9-]+$/.test(set)) throw new Error(`set name: [a-z0-9-], not ${set}`)
  const built = ((await import(`./lines/${set}.js`)) as LinesSet).buildLinesCases()
  const files = new Map(await Promise.all(built.cases.map(async (c) => [c.id, (await writeXlsx(built.books[c.book])).data] as const)))
  const sent = Object.fromEntries(await Promise.all(built.cases.map(async (c) => [c.id, await partsHash(files.get(c.id)!)] as const)))
  return { built, files, sent }
}

/** A step as the AppleScript takes it: sheet, range, number format, bold, the fill's [r, g, b], the value's kind and text. */
export const stepArgs = (c: LinesCase) =>
  (c.steps ?? []).map((s) => [
    s.sheet,
    s.range,
    s.format ?? '',
    !!s.bold,
    s.fill ? [1, 3, 5].map((i) => parseInt(s.fill!.slice(i, i + 2), 16)) : [],
    s.value === undefined ? '' : typeof s.value === 'number' ? 'n' : 's',
    s.value === undefined ? '' : String(s.value),
  ])

type RawLook = [address: string, text: string, format: string, fill: number[], colorIndex: string, font: number[], bold: boolean, italic: boolean]

/** One case in Excel: its answer and the saved file, or the repair window's text. */
async function runCase(c: LinesCase, data: Uint8Array): Promise<{ answer: LinesAnswer; saved: Uint8Array } | { rejected: string }> {
  const r = await runSavingScript(c.id, SCRIPT, data, [JSON.stringify(stepArgs(c)), JSON.stringify(c.read)])
  if ('rejected' in r) return r
  const raw = JSON.parse(r.out) as { error: string; looks: [string, RawLook[]][] }
  const looks: LinesAnswer['looks'] = {}
  for (const [sheet, cells] of raw.looks)
    looks[sheet] = Object.fromEntries(
      cells.map(([a, text, format, fill, ci, font, bold, italic]): [string, LineLook] => [a.replaceAll('$', ''), { text, format, fill: ci === 'color index none' ? null : fromRgb(fill), font: fromRgb(font), bold, italic }]),
    )
  return { answer: { ...(raw.error ? { error: raw.error } : {}), looks }, saved: r.saved }
}

/** Runs the set's cases that have no stored answer for the same case and file (FCOV_FRESH=1 runs all) and stores the answers. */
export async function stepLines(set: string): Promise<void> {
  const { built, files, sent } = await linesSet(set)
  const target = join(ROOT, `test/excel/lines/${set}.json`)
  const dir = join(ROOT, `test/excel/lines/${set}`)
  const prev = existsSync(target) && !process.env.FCOV_FRESH ? (JSON.parse(readFileSync(target, 'utf8')) as StoredLines) : null
  const { answers, rejected } = keptAnswers(built.cases, sent, prev, (id) => existsSync(join(dir, `${id}.xlsx`)))
  mkdirSync(dir, { recursive: true })
  const t0 = Date.now()
  const save = () => {
    const meta = { version: excelVersion(), locale: 'en-US (-AppleLanguages (en) -AppleLocale en_US)', seconds: Math.round((Date.now() - t0) / 100) / 10 }
    const refused = inCaseOrder(built.cases, rejected)
    writeFileSync(target, JSON.stringify({ meta, ...built, sent: inCaseOrder(built.cases, sent), answers: inCaseOrder(built.cases, answers), ...(Object.keys(refused).length ? { rejected: refused } : {}) }, null, 1) + '\n')
  }
  await ensureExcel('en-US')
  let done = 0
  try {
    const only = process.env.FLINES_ONLY?.split(',')
    for (const c of built.cases) {
      if (answers[c.id] || rejected[c.id] || (only && !only.includes(c.id))) continue
      const r = await runCase(c, files.get(c.id)!)
      if ('rejected' in r) {
        rejected[c.id] = r.rejected
        console.log(`${c.id}: ${r.rejected}`)
      } else {
        answers[c.id] = r.answer
        writeFileSync(join(dir, `${c.id}.xlsx`), r.saved)
      }
      done++
      save()
    }
  } finally {
    save()
    console.log(`lines ${set}: ${done} cases run, ${Object.keys(answers).length} of ${built.cases.length} answered, ${Object.keys(rejected).length} not loaded, ${((Date.now() - t0) / 1000).toFixed(1)} s`)
  }
}
