/**
 * Formats of rows and columns against Excel 16.113 for Mac and Excel 16.0.20430 for Windows, en-US (`test/excel/lines/`,
 * `test/excel/win/lines/`, from `run.ts lines <set>` and `bench/windows/run.ts lines <set>`): what the read cells show
 * after each case (`Range.Text`, the number format, the display format's fill and font) against the engine on the book
 * written (on the book the reader reads for a bare case), what the engine shows on the book read from the file Excel
 * saved, a book written and saved by Excel read back as written, and how Excel stores the cells of formatted lines.
 * Known differences are pinned with their reasons.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredLines } from '../bench/coverage/lines.js'
import { caseBook, caseFile, engineCase, partsOf, sessionBook, sessionCase, type LineLook, type LinesAnswer } from '../bench/coverage/lines/cases.js'
import { compareBooks, entries } from '../bench/coverage/xlsx/compare.js'
import { partsHash } from '../bench/coverage/xlsx/sent.js'
import type { BookInput, NumberFormat } from '../src/contract.js'
import { checkBook } from '../src/input.js'
import { readXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'

/** Excel merges the styles of a row and a column it formats across each other; `applySteps` sets the line's style, as `setLayout` does. */
const MERGED =
  "Excel gives the cells where a line it formats crosses a formatted line the two lines' styles merged field by field (the row's bold or italic, the column's fill); the step here sets the line's style alone, as setLayout does, and the model shows such a cell the row's style, a style being one whole"

/** Differences pinned by `<set>:<case>:<sheet>!<cell>`, on both Excels. */
const LINES_KNOWN: Record<string, string> = {
  'files:excel-row-column:S!H8': MERGED,
  'files:excel-row-column:S!I8': MERGED,
  'held:excel-italic-fill:H!E7': MERGED,
}

type LinesSetFile = StoredLines & { name: string; dir: URL }

function linesSets(sub: string): LinesSetFile[] {
  const dir = new URL(`./excel/${sub}`, import.meta.url)
  if (!existsSync(dir)) return []
  return readdirSync(dir)
    .filter((f) => f.endsWith('.json'))
    .sort()
    .map((f) => ({ name: f.slice(0, -5), dir: new URL(`${f.slice(0, -5)}/`, dir), ...(JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as StoredLines) }))
}

const show = (l: LineLook | undefined) => (l ? `${JSON.stringify(l.text)} ${l.format} fill ${l.fill ?? 'none'} font ${l.font}${l.bold ? ' bold' : ''}${l.italic ? ' italic' : ''}` : 'none')

/** A format code without Excel's escapes of literal characters (`dddd\\,\\ mmmm` is `dddd, mmmm`), which it saves and does not show. */
const plainCode = <T>(f: T): T => (typeof f === 'string' ? (f.replace(/\\(.)/g, '$1') as T) : f)

const key = (l: LineLook | undefined) => l && JSON.stringify([l.text, plainCode(l.format), l.fill, l.font, l.bold, l.italic])

/** A book with its format codes without Excel's escapes. */
function unescaped(book: BookInput): BookInput {
  const lines = <T extends { numberFormat?: NumberFormat }>(m: Record<string, T> | undefined) => m && Object.fromEntries(Object.entries(m).map(([k, v]) => [k, v.numberFormat === undefined ? v : { ...v, numberFormat: plainCode(v.numberFormat) }]))
  return { ...book, sheets: book.sheets.map((s) => ({ ...s, cells: lines(s.cells)!, columns: lines(s.columns), rows: lines(s.rows) })) }
}

/** The cells whose look differs: `<sheet>!<cell>: Excel …, engine …`. */
function differences(excel: LinesAnswer['looks'], engine: LinesAnswer['looks']): Map<string, string> {
  const out = new Map<string, string>()
  for (const [sheet, cells] of Object.entries(excel))
    for (const [addr, l] of Object.entries(cells)) if (key(l) !== key(engine[sheet]?.[addr])) out.set(`${sheet}!${addr}`, `Excel ${show(l)}, engine ${show(engine[sheet]?.[addr])}`)
  return out
}

const saved = async (set: LinesSetFile, id: string) => readXlsx(new Uint8Array(readFileSync(new URL(`${id}.xlsx`, set.dir))))

const PLATFORMS: [string, LinesSetFile[]][] = [
  ['Excel for Mac', linesSets('lines/')],
  ['Excel for Windows', linesSets('win/lines/')],
]

describe.each(PLATFORMS)('formats of rows and columns against %s', (_, sets) => {
  it('every set is answered', () => expect(sets.map((s) => s.name).filter((n) => n !== 'clear' && n !== 'patch')).toEqual(['files', 'held']))
  it('every pinned difference names a case', () => {
    const ids = new Set(sets.flatMap((s) => s.cases.map((c) => `${s.name}:${c.id}`)))
    expect(Object.keys(LINES_KNOWN).filter((k) => !ids.has(k.split(':').slice(0, 2).join(':')))).toEqual([])
  })

  for (const set of sets)
    describe(set.name, () => {
      it('Excel loaded every book, sent as the writer writes it now', async () => {
        expect(set.rejected ?? {}).toEqual({})
        for (const c of set.cases) expect(set.sent[c.id], c.id).toBe(await partsHash(await caseFile(set, c)))
      })
      for (const c of set.cases) {
        const a = set.answers[c.id]
        const pinned = (cell: string) => LINES_KNOWN[`${set.name}:${c.id}:${cell}`]
        // a set of a session's steps (clearing among them) has no answer of the model's
        it.skipIf(set.session)(`${c.id}: the cells show what Excel showed`, async () => {
          expect(a.error).toBeUndefined()
          const got = differences(a.looks, engineCase(c, await caseBook(set, c)))
          expect([...got].filter(([cell]) => !pinned(cell)).map(([cell, d]) => `${cell}: ${d}`)).toEqual([])
          for (const k of Object.keys(LINES_KNOWN).filter((x) => x.startsWith(`${set.name}:${c.id}:`))) expect(got.has(k.split(':')[2]), `${k} pinned`).toBe(true)
        })
        it(`${c.id}: the steps taken by a session (patchCells, setCells, clearCells), the cells show what Excel showed, the lines Excel formats across each other too`, async () => {
          expect([...differences(a.looks, sessionCase(c, await caseBook(set, c)))].map(([cell, d]) => `${cell}: ${d}`)).toEqual([])
        })
        if (c.parts)
          it(`${c.id}: the notes, hyperlinks, conditional formats and rules of data validation of Excel's file are the session's`, async () => {
            expect(partsOf((await saved(set, c.id)).book)).toEqual(partsOf(sessionBook(c, await caseBook(set, c))))
          })
        it(`${c.id}: on the book read from Excel's file the engine shows what Excel showed`, async () => {
          const r = await saved(set, c.id)
          expect(r.issues).toEqual([])
          expect([...differences(a.looks, engineCase({ ...c, steps: undefined }, r.book))].map(([cell, d]) => `${cell}: ${d}`)).toEqual([])
        })
        if (!c.steps && !c.bare)
          it(`${c.id}: Excel's file reads back as the book written`, async () => {
            expect(compareBooks(entries(checkBook(set.books[c.book])), entries(unescaped((await saved(set, c.id)).book)))).toEqual([])
          })
      }
    })
})

describe('how Excel stores the cells of formatted lines (the saved files of lines/files)', () => {
  const [mac, win] = PLATFORMS.map(([, sets]) => sets.find((s) => s.name === 'files')!)
  const sheetXml = async (set: LinesSetFile, id: string) => (await openArchive(new Uint8Array(readFileSync(new URL(`${id}.xlsx`, set.dir)))).text('xl/worksheets/sheet1.xml'))!
  const cellXf = (xml: string, addr: string) => [...xml.matchAll(/<c r="([A-Z]+\d+)"(?: s="(\d+)")?/g)].find((m) => m[1] === addr)?.[2] ?? '0'

  it.each([mac, win])('a value typed into a formatted column or row is stored with the line\'s cell format', async (set) => {
    const column = await sheetXml(set, 'excel-column')
    const style = /<col min="2" max="2"[^>]* style="(\d+)"/.exec(column)![1]
    expect([cellXf(column, 'B2'), cellXf(column, 'B3')]).toEqual([style, style])
    const row = await sheetXml(set, 'excel-row')
    const s = /<row r="4"[^>]* s="(\d+)" customFormat="1"/.exec(row)![1]
    expect([cellXf(row, 'D4'), cellXf(row, 'E4')]).toEqual([s, s])
  })

  it.each([mac, win])('Excel writes the empty cells where lines it formats cross, with both lines\' formats', async (set) => {
    for (const [id, empty] of [['excel-column-row', 'G6'], ['excel-row-column', 'I8'], ['excel-format-fill', 'K10']])
      expect(await sheetXml(set, id), `${id} ${empty}`).toMatch(new RegExp(`<c r="${empty}" s="\\d+"/>`))
  })
})
