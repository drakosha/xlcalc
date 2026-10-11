/**
 * Notes, hyperlinks and frozen panes in Excel for Windows: each case of `bench/coverage/parts/<set>.ts` in the file the
 * Mac step sends (`writeXlsx`, threads of comments put in by hand), what it makes through COM (`AddComment`,
 * `Hyperlinks.Add`, `FreezePanes`), its operation, every sheet's parts read (threads of comments too, `CommentsThreaded`,
 * which Excel 2024 does not list) and the book saved. Answers go to `test/excel/win/parts/<set>.json` in the Mac's shape, the saved files beside them;
 * `test/sheet-parts-win.test.ts` reads them.
 */
import { partsSet } from '../coverage/parts.js'
import type { PartsAnswer, PartsCase, PartsMake } from '../coverage/parts/cases.js'
import { fileAnswers, secondsSince } from '../coverage/stored.js'
import { fileJobs, makeErrors, opSteps } from './books.js'
import { winFile } from './functions.js'
import { metaOf, type Step, type Tag, type WinEnv } from './remote.js'

/** A step of `xl.ps1` that makes what the case makes. */
function makeStep(m: PartsMake): Step {
  if ('note' in m) return { do: 'make', kind: 'note', sheet: m.sheet, cell: m.note, text: m.text }
  if ('link' in m) return { do: 'make', kind: 'link', sheet: m.sheet, cell: m.link, url: m.url ?? '', location: m.location ?? '', tip: m.tooltip ?? '' }
  return { do: 'make', kind: 'freeze', sheet: m.freeze, rows: m.rows, columns: m.columns }
}

type RawSheet = [string, [string, string, string][], [string, [string, string][]][], [string, string, string, string, string][], [boolean, number, number, number, number, number, number]]

const plain = (a: string) => a.replaceAll('$', '')

export async function stepParts(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const { built, files, sent } = await partsSet(set)
  const { books, cases, threads } = built
  const store = fileAnswers<PartsCase, PartsAnswer>(winFile(`parts/${set}.json`), cases, sent)
  const t0 = Date.now()
  const save = () => store.save({ ...metaOf(e, tag), seconds: secondsSince(t0) }, { books, cases, ...(threads ? { threads } : {}) })
  await fileJobs(
    `parts ${set}`,
    `parts-${set}`,
    store,
    files,
    (c) => [...(c.make ?? []).map(makeStep), ...opSteps(c), { do: 'parts' }, { do: 'saveAs' }],
    (c, a) => {
      const { n, errors } = makeErrors(c, a.answers)
      const raw = a.answers[n] as RawSheet[]
      return {
        ...(errors.length ? { error: errors.join('; ') } : {}),
        sheets: raw.map(([name, notes, threads, links, [freeze, splitRow, splitColumn, , , scrollRow, scrollColumn]]) => ({
          name,
          notes: Object.fromEntries(notes.map(([addr, author, text]) => [plain(addr), { author, text }])),
          links: links.map(([range, address, subAddress, tip, display]) => ({ range: plain(range), address, subAddress, tip, display })),
          panes: { freeze, splitRow, splitColumn, scrollRow, scrollColumn },
          ...(threads.length ? { threads: Object.fromEntries(threads.map(([addr, list]) => [plain(addr), list.map(([author, text]) => ({ author, text }))])) } : {}),
        })),
      }
    },
    save,
  )
}
