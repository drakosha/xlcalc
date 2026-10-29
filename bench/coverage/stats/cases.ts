/**
 * Excel's Workbook Statistics (`run.ts stats <set>`): a set names its cases, each a book written by `writeXlsx` (with
 * threads of comments put in, `parts/threaded.ts`) and the sheet active when Excel opens Review > Workbook Statistics.
 * The answer is what its window shows, row by row, for the sheet and for the workbook; `predict.ts` writes what
 * `Session.getWorkbookStatistics` gives before Excel runs a set, as Excel would show it.
 */
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { BookInput, StatisticsCounts, WorkbookStatistics } from '../../../src/contract.js'
import { createSession } from '../../../src/index.js'
import { readXlsx, writeXlsx } from '../../../src/xlsx.js'
import { threadedFile } from '../parts/threaded.js'
import type { ThreadSpec } from '../parts/cases.js'

const ROOT = join(fileURLToPath(new URL('.', import.meta.url)), '../../..')

export interface StatsCase {
  id: string
  book: BookInput
  /** The sheet active when the window opens. */
  sheet: string
  threads?: ThreadSpec[]
}

/** A row of the window: its label and the text it shows. */
export type StatsRow = [label: string, text: string]

/** What the window showed: the rows of the current sheet and of the workbook, in its order. */
export interface StatsAnswer {
  sheet: StatsRow[]
  workbook: StatsRow[]
}

export interface StatsSet {
  buildStatsCases(): StatsCase[]
}

export const statsFile = (set: string) => join(ROOT, 'test/excel/stats', `${set}.json`)

/** A case's file: written by `writeXlsx`, with its threads of comments put in. */
export async function statsFileOf(c: StatsCase): Promise<Uint8Array> {
  return c.threads ? threadedFile(c.book, c.threads) : (await writeXlsx({ book: c.book })).data
}

/** The window's labels of each count, in the order Excel for Windows lists them. */
const COUNTS = [['Cells with data', 'cells'], ['Tables', 'tables'], ['PivotTables', 'pivotTables'], ['Formulas', 'formulas'], ['Charts', 'charts'], ['Images', 'images'], ['Form Controls', 'formControls'], ['Objects', 'objects'], ['Comments', 'comments'], ['Notes', 'notes']] as const

/** The rows of the engine's statistics as the window labels them, a count the workbook does not hold as `null`. */
export function statsRows(st: WorkbookStatistics): StatsAnswer {
  const rows = (x: StatisticsCounts) => COUNTS.map(([label, k]): StatsRow => [label, String(x[k])])
  return { sheet: [['End of sheet', st.sheet.end], ...rows(st.sheet)], workbook: [['Sheets', String(st.workbook.sheets)], ...rows(st.workbook)] }
}

/** The engine's statistics of a case: over the book the reader reads from the case's file, with its sheet active. */
export async function engineStats(c: StatsCase): Promise<StatsAnswer> {
  const s = createSession((await readXlsx(await statsFileOf(c))).book)
  const out = statsRows(s.getWorkbookStatistics(c.sheet))
  s.close()
  return out
}

/**
 * The lines where the window and the engine differ: a row Excel shows with another value or that the engine has no
 * count for, and a count of the engine Excel does not show though it is not 0 (Excel for Mac leaves out rows of none).
 */
export function statsDifferences(excel: StatsAnswer, engine: StatsAnswer): string[] {
  const out: string[] = []
  for (const part of ['sheet', 'workbook'] as const) {
    const mine = new Map(engine[part])
    const shown = new Set(excel[part].map(([label]) => label))
    for (const [label, text] of excel[part]) if (mine.get(label) !== text) out.push(`${part} ${label}: Excel ${text}, engine ${mine.get(label) ?? 'none'}`)
    for (const [label, text] of engine[part]) if (!shown.has(label) && text !== '0' && text !== 'null') out.push(`${part} ${label}: not shown, engine ${text}`)
  }
  return out
}
