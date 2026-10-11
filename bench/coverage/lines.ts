/**
 * Formats of rows and columns against Excel for Mac en-US: `npx tsx bench/coverage/run.ts lines <set>` writes each book
 * of `lines/<set>.ts` with `writeXlsx`, Excel opens it, takes the case's steps, computes the book, reads what the listed
 * cells show and their number format, and saves the book (`excel-lines.applescript`). The answers go to
 * `test/excel/lines/<set>.json` and the saved files to `test/excel/lines/<set>/<id>.xlsx`; `test/lines-excel.test.ts`
 * compares them with the engine and the reader without Excel. `FLINES_ONLY=id,id` runs only those cases,
 * `FCOV_FRESH=1` every case again.
 */
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { fromRgb } from './cond/cases.js'
import { ensureExcel, excelVersion, runSavingScript } from './excel.js'
import { caseFile, type LineLook, type LinesAnswer, type LinesCase, type LinesCases, type LinesSet } from './lines/cases.js'
import { localeMeta } from './locales.js'
import { fileAnswers, runFileCases, secondsSince, setModule } from './stored.js'
import { caseFiles } from './xlsx/sent.js'

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
  const built = (await setModule<LinesSet>('lines', set)).buildLinesCases()
  return { built, ...(await caseFiles(built.cases, (c) => caseFile(built, c))) }
}

/** A step as the AppleScript takes it: sheet, range, number format, bold, italic, the fill's [r, g, b], the value's kind and text, what it clears. */
export const stepArgs = (c: LinesCase) =>
  (c.steps ?? []).map((s) => [
    s.sheet,
    s.range,
    s.format ?? '',
    !!s.bold,
    !!s.italic,
    s.fill ? [1, 3, 5].map((i) => parseInt(s.fill!.slice(i, i + 2), 16)) : [],
    s.value === undefined ? '' : typeof s.value === 'number' ? 'n' : 's',
    s.value === undefined ? '' : String(s.value),
    s.clear ?? '',
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
  const store = fileAnswers<LinesCase, LinesAnswer>(join(ROOT, `test/excel/lines/${set}.json`), built.cases, sent)
  const t0 = Date.now()
  const save = () => store.save({ version: excelVersion(), locale: localeMeta('en-US'), seconds: secondsSince(t0) }, built)
  await ensureExcel('en-US')
  await runFileCases(`lines ${set}`, store, store.todo, 'FLINES_ONLY', (c) => runCase(c, files.get(c.id)!), save, t0)
}
