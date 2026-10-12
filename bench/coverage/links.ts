/**
 * Links to other workbooks against Excel en-US: `npx tsx bench/coverage/run.ts links <set>` writes each case of
 * `links/<set>.ts` (`linkFile`), Excel opens it without updating its links (`update links do not update links`), lists
 * the linked books, reads the rectangles after opening, takes the case's operations, reads them again after a full
 * calculation, and saves the book
 * (`excel-links.applescript`). The answers go to `test/excel/links/<set>.json`, the saved files to
 * `test/excel/links/<set>/<id>.xlsx`. A window Excel shows meanwhile (a question about the links) stops the run with its
 * text, nothing clicked. `FLINKS_ONLY=id,id` runs only those cases, `FCOV_FRESH=1` every case again.
 */
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { answerOf, cellsOf, eachRectCell, grid, mirrorRanges, refusedOf } from './case-book.js'
import { ensureExcel, excelVersion, runSavingScript } from './excel.js'
import { linkFile, type LinksAnswer, type LinksCase, type LinksCases, type LinksOp, type LinksRead, type LinksSet } from './links/cases.js'
import { localeMeta } from './locales.js'
import { fileAnswers, runFileCases, secondsSince, setModule } from './stored.js'
import { caseFiles } from './xlsx/sent.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
const SCRIPT = join(HERE, 'excel-links.applescript')

export interface StoredLinks extends LinksCases {
  meta: Record<string, string | number>
  sent: Record<string, string>
  answers: Record<string, LinksAnswer>
  rejected?: Record<string, string>
}

export const linksFile = (set: string) => join(ROOT, `test/excel/links/${set}.json`)

/** The set's cases, the file of each and what identifies it. */
export async function linksSet(set: string): Promise<{ built: LinksCases; files: Map<string, Uint8Array>; sent: Record<string, string> }> {
  const built = (await setModule<LinksSet>('links', set)).buildLinksCases()
  return { built, ...(await caseFiles(built.cases, linkFile)) }
}

/** One read of the script's output as an answer's read. */
function readOf(c: LinksCase, snap: unknown[][]): LinksRead {
  const { values, formulas } = answerOf(c.read, snap.map((x) => x.slice(0, 4)) as [unknown, unknown, unknown, unknown][])
  const local: Record<string, string> = {}
  const shown: Record<string, string> = {}
  eachRectCell(c.read, (key, i, r, j) => {
    const fl = String(grid(snap[i][4])[r][j])
    if (fl.startsWith('=')) local[key] = fl
    shown[key] = String((snap[i][5] as string[])[r * grid(snap[i][4])[0].length + j])
  })
  return { values, formulas, local, shown }
}

/** An operation as a line of the script's argument: op, sheet, range, to. */
const opLine = (o: LinksOp) => [o.op, o.sheet, 'range' in o ? o.range : '', 'to' in o ? o.to : ''].join('\t')

async function runCase(c: LinksCase, data: Uint8Array): Promise<{ answer: LinksAnswer; saved: Uint8Array } | { rejected: string }> {
  const at = mirrorRanges(c.read)
  const rects = c.read.map((g, i) => `${g.sheet}\t${g.range}\t${at[i]}\t${cellsOf(g.range).addrs.join(',')}`).join('\n')
  const r = await runSavingScript(c.id, SCRIPT, data, [rects, (c.ops ?? []).map(opLine).join('\n')], 120_000)
  if ('rejected' in r) return r
  const raw = JSON.parse(r.out) as { sources: string[]; folder: string; errors: string[]; reads: unknown[][][] }
  const errors = refusedOf(raw.errors)
  return { answer: { sources: raw.sources, folder: raw.folder, ...(errors ? { errors } : {}), reads: raw.reads.map((snap) => readOf(c, snap)) }, saved: r.saved }
}

export async function stepLinks(set: string): Promise<void> {
  const { built, files, sent } = await linksSet(set)
  const store = fileAnswers<LinksCase, LinksAnswer>(linksFile(set), built.cases, sent)
  const t0 = Date.now()
  const save = () => store.save({ version: excelVersion(), locale: localeMeta('en-US'), seconds: secondsSince(t0) }, built)
  await ensureExcel('en-US')
  await runFileCases(`links ${set}`, store, store.todo, 'FLINKS_ONLY', (c) => runCase(c, files.get(c.id)!), save, t0)
}
