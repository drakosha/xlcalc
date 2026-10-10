/**
 * Notes, hyperlinks and frozen panes in Excel for Windows: each case of `bench/coverage/parts/<set>.ts` in the file the
 * Mac step sends (`writeXlsx`), what it makes through COM (`AddComment`, `Hyperlinks.Add`, `FreezePanes`, and threads of
 * comments, `AddCommentThreaded` with `AddReply`, which only Windows has), its operation, every sheet's parts read and the
 * book saved. Answers go to `test/excel/win/parts/<set>.json` in the Mac's shape, the saved files beside them;
 * `test/sheet-parts-win.test.ts` reads them.
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { partsFile, type StoredParts } from '../coverage/parts.js'
import type { PartsAnswer, PartsCase, PartsMake, PartsSet } from '../coverage/parts/cases.js'
import { inCaseOrder } from '../coverage/stored.js'
import { opArgs } from '../coverage/struct.js'
import { partsHash } from '../coverage/xlsx/sent.js'
import { caseJobs } from './books.js'
import { winFile } from './functions.js'
import { metaOf, type Step, type Tag, type WinEnv } from './remote.js'

/** A step of `xl.ps1` that makes what the case makes. */
function makeStep(m: PartsMake): Step {
  if ('note' in m) return { do: 'make', kind: 'note', sheet: m.sheet, cell: m.note, text: m.text }
  if ('link' in m) return { do: 'make', kind: 'link', sheet: m.sheet, cell: m.link, url: m.url ?? '', location: m.location ?? '', tip: m.tooltip ?? '' }
  if ('freeze' in m) return { do: 'make', kind: 'freeze', sheet: m.freeze, rows: m.rows, columns: m.columns }
  return { do: 'make', kind: 'thread', sheet: m.sheet, cell: m.thread, texts: m.texts }
}

type RawSheet = [string, [string, string, string][], [string, [string, string][]][], [string, string, string, string, string][], [boolean, number, number, number, number, number, number]]

const plain = (a: string) => a.replaceAll('$', '')

export async function stepParts(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const { books, cases } = ((await import(`../coverage/parts/${set}.js`)) as PartsSet).buildPartsCases()
  const file = winFile(`parts/${set}.json`)
  const dir = winFile(`parts/${set}`)
  const prev = existsSync(file) && !process.env.FCOV_FRESH ? (JSON.parse(readFileSync(file, 'utf8')) as StoredParts) : null
  const files = new Map(await Promise.all(cases.map(async (c) => [c.id, await partsFile(c, books)] as const)))
  const sent = Object.fromEntries(await Promise.all(cases.map(async (c) => [c.id, await partsHash(files.get(c.id)!)] as const)))
  const key = (c: PartsCase, hash: string | undefined) => JSON.stringify([c, hash])
  const same = new Map(prev?.cases.map((c) => [c.id, key(c, prev.sent[c.id])]) ?? [])
  const answers: Record<string, PartsAnswer> = {}
  const rejected: Record<string, string> = {}
  for (const c of cases)
    if (same.get(c.id) === key(c, sent[c.id])) {
      if (prev!.answers[c.id] && existsSync(join(dir, `${c.id}.xlsx`))) answers[c.id] = prev!.answers[c.id]
      else if (prev!.rejected?.[c.id]) rejected[c.id] = prev!.rejected[c.id]
    }
  mkdirSync(dir, { recursive: true })
  const t0 = Date.now()
  const save = () => {
    const refused = inCaseOrder(cases, rejected)
    mkdirSync(dirname(file), { recursive: true })
    writeFileSync(file, JSON.stringify({ meta: { ...metaOf(e, tag), seconds: Math.round((Date.now() - t0) / 100) / 10 }, books, cases, sent: inCaseOrder(cases, sent), answers: inCaseOrder(cases, answers), ...(Object.keys(refused).length ? { rejected: refused } : {}) }, null, 1) + '\n')
  }
  const todo = cases.filter((c) => !answers[c.id] && !rejected[c.id])
  await caseJobs(
    `parts-${set}`,
    todo,
    async (c, f) => {
      writeFileSync(f, files.get(c.id)!)
      const [sheet, range, target, to] = c.op ? opArgs({ id: c.id, book: c.book, op: c.op }) : []
      const op: Step[] = c.op ? [{ do: 'struct', op: c.op.op, sheet, range, target, to }] : []
      return { file: f, steps: [...(c.make ?? []).map(makeStep), ...op, { do: 'parts' }, { do: 'saveAs' }] }
    },
    (c, a) => {
      const n = (c.make ?? []).length + (c.op ? 1 : 0)
      const errors = (a.answers.slice(0, n) as ({ error?: string } | null)[]).flatMap((x, i) => (x?.error ? [`${i < (c.make ?? []).length ? 'make' : 'op'}: ${x.error}`] : []))
      const raw = a.answers[n] as RawSheet[]
      answers[c.id] = {
        ...(errors.length ? { error: errors.join('; ') } : {}),
        sheets: raw.map(([name, notes, threads, links, [freeze, splitRow, splitColumn, , , scrollRow, scrollColumn]]) => ({
          name,
          notes: Object.fromEntries(notes.map(([addr, author, text]) => [plain(addr), { author, text }])),
          links: links.map(([range, address, subAddress, tip, display]) => ({ range: plain(range), address, subAddress, tip, display })),
          panes: { freeze, splitRow, splitColumn, scrollRow, scrollColumn },
          ...(threads.length ? { threads: Object.fromEntries(threads.map(([addr, list]) => [plain(addr), list.map(([author, text]) => ({ author, text }))])) } : {}),
        })),
      }
      if (a.saved) writeFileSync(join(dir, `${c.id}.xlsx`), a.saved)
      else {
        delete answers[c.id]
        rejected[c.id] = 'Excel did not save the book'
      }
    },
    (c, why) => {
      console.log(`${c.id}: Excel did not load the book (${why})`)
      rejected[c.id] = why
    },
    save,
  )
  save()
  console.log(`parts ${set}: ${todo.length} cases run, ${Object.keys(answers).length} of ${cases.length} answered, ${Object.keys(rejected).length} not loaded`)
}
