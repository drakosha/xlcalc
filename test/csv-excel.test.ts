/**
 * CSV files against Excel for Mac (`test/excel/csv/`, from the step `csv`): `readCsv` read as Excel opens a file (the
 * locale's list separator, formulas taken, the year of the run's day) gives each cell the value or formula, the number
 * format and the wrapping that the xlsx Excel saved after opening it holds, and the sheet the name Excel gave it;
 * `writeCsv` of the book Excel read, or of a case's book, writes the CSV Excel saved. The sets `held` and `held-ru` were
 * predicted (`bench/coverage/csv/<set>.predicted.json`) before Excel ran them. Where Excel for Mac reads or writes
 * otherwise, by a reason, the cases are listed: the engine follows Excel for Windows' code pages and built-in formats.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { csvFiles } from '../bench/coverage/csv.js'
import { engineAnswer, type CsvAnswer, type CsvCase, type CsvSet } from '../bench/coverage/csv/cases.js'
import type { BookInput, CellInput } from '../src/contract.js'
import { registerLocale } from '../src/index.js'
import { ruRU } from '../src/locales/ru-RU.js'
import { readXlsx } from '../src/xlsx.js'
import { localeSets } from './stored-sets.js'

registerLocale(ruRU)

interface StoredCsv {
  meta: { date: string }
  cases: CsvCase[]
  sent: Record<string, string>
  answers: Record<string, CsvAnswer>
}

/** Excel for Mac reads a file without a byte order mark in its Mac code page (Mac Roman, Mac Cyrillic), the engine in Windows': text outside ASCII differs. */
const MAC_PAGE = 'Excel for Mac reads a file without a byte order mark in Mac Roman or Mac Cyrillic, the engine as UTF-8 or in the Windows code page'
/** Excel for Mac reads no UTF-16BE and no UTF-16 without a mark. */
const MAC_UTF16 = 'Excel for Mac reads UTF-16BE and UTF-16 without a mark as single bytes'
/** Typed errors the entry does not take as errors. */
const ERRORS = 'Excel takes #SPILL!, #CALC! and #GETTING_DATA (as #N/A) typed as errors; the entry reads them as text'
/** Money the entry gives no format. */
const EURO = 'Excel gives 5 €, €-5 and -€5 euro formats; the entry reads the number without one'

/** The cells where the engine's book knowingly differs from Excel for Mac's, by set and case. */
const KNOWN_READ: Record<string, Record<string, [string, string[]]>> = {
  files: {
    'numbers-more': [MAC_PAGE, ['M1', 'N1']],
    utf8: [MAC_PAGE, ['A1', 'B1', 'C1', 'D1']],
    cp1252: [MAC_PAGE, ['A1', 'B1', 'C1', 'D1']],
    cp1251: [MAC_PAGE, ['A1', 'B1']],
    'utf16be-bom': [MAC_UTF16, ['A1', 'B1', 'A2', 'B2']],
    utf16le: [MAC_UTF16, ['A2', 'B2']],
  },
  held: { words: [ERRORS, ['F1', 'G1', 'H1']], 'euro-bom': [EURO, ['B1', 'C1', 'D1']] },
  'files-ru': {
    basic: [MAC_PAGE, ['A1', 'B1', 'C1', 'A2', 'A3']],
    money: [MAC_PAGE, ['A1', 'B1', 'C1', 'D1', 'F1', 'G1']],
    dates: [MAC_PAGE, ['F1']],
    logic: [MAC_PAGE, ['A1', 'B1', 'D1', 'E1', 'G1']],
    formulas: [MAC_PAGE, ['B1', 'C1', 'G1']],
    utf8: [MAC_PAGE, ['A1', 'B1', 'C1']],
    cp1251: [MAC_PAGE, ['A1', 'C1']],
  },
  'held-ru': { money: [EURO, ['I1']] },
}

/** Excel for Mac saves CSV in its Mac code page; it shows the built-in date format and month names otherwise. */
const MAC_SAVE = 'Excel for Mac saves CSV in Mac Roman or Mac Cyrillic, the engine in the Windows code page'
const MAC_DATE = 'Excel for Mac shows the built-in date format 14 as m/d/yy in en-US; the engine, as Excel for Windows, m/d/yyyy'
const MAC_NAMES = 'Excel for Mac names the months in ru-RU otherwise (март, февр, its time 01:30:45); the engine, as Excel for Windows, мар, фев, Март, 1:30:45'
const MAC_FORMULA = 'in a ru-RU book the engine computes =SUM(1) typed with the English name; Excel gives #NAME?'

/** The cases whose CSV knowingly differs from what Excel for Mac saved (besides those whose book already differs). */
const KNOWN_WRITE: Record<string, Record<string, string>> = {
  files: { dates: MAC_DATE, 'w-formats': MAC_DATE, 'utf8-bom': MAC_SAVE, 'utf16le-bom': MAC_SAVE, 'w-unicode': MAC_SAVE },
  held: { dates: MAC_DATE },
  'files-ru': { numbers: MAC_NAMES, 'utf8-bom': MAC_SAVE, 'w-formats': MAC_SAVE, 'w-unicode': MAC_SAVE },
  'held-ru': { dates: MAC_NAMES, formulas: MAC_FORMULA, 'w-formats': MAC_NAMES },
}

/** A cell as both sides are compared: its value or formula and number format, and whether it wraps. */
function looks(book: BookInput): Record<string, object> {
  const out: Record<string, object> = {}
  for (const [a, { style, ...c }] of Object.entries(book.sheets[0].cells) as [string, CellInput][]) {
    const wrap = style !== undefined && book.styles?.[style]?.alignment?.wrap === true
    out[a] = { ...(c.formula !== undefined ? { formula: c.formula } : { value: c.value }), ...(c.numberFormat !== undefined ? { numberFormat: c.numberFormat } : {}), ...(wrap ? { wrap } : {}) }
  }
  return out
}

/** The cells of `a` that differ from `b`'s, either way. */
const differing = (a: Record<string, object>, b: Record<string, object>) => [...new Set([...Object.keys(a), ...Object.keys(b)])].filter((k) => JSON.stringify(a[k]) !== JSON.stringify(b[k])).sort()

const sets = localeSets<StoredCsv>('csv')

describe.each(sets)('CSV against Excel for Mac, $tag $set', ({ tag, set, stored }) => {
  const now = `${stored.meta.date}T12:00:00Z`
  const built = async () => csvFiles(((await import(`../bench/coverage/csv/${set}.ts`)) as CsvSet).buildCsvCases())
  const dir = new URL(`./excel/csv/${tag === 'en-US' ? '' : `${tag}/`}${set}/`, import.meta.url)

  it('the files are the ones Excel opened', async () => {
    expect((await built()).sent).toEqual(stored.sent)
  })

  it('readCsv reads each file as Excel opened it, but where the platforms part', async () => {
    const { files } = await built()
    const known = KNOWN_READ[set] ?? {}
    for (const c of stored.cases.filter((x) => !x.book)) {
      const engine = engineAnswer(c, files.get(c.id)!, tag, now)
      const excel = (await readXlsx(readFileSync(new URL(`${c.id}.xlsx`, dir)))).book
      expect(engine.sheet, c.id).toBe(stored.answers[c.id].sheet)
      const ours = looks({ sheets: [{ name: engine.sheet!, cells: engine.cells! }], styles: engine.styles })
      expect(differing(ours, looks(excel)), `${c.id}: ${known[c.id]?.[0] ?? ''}`).toEqual((known[c.id]?.[1] ?? []).slice().sort())
    }
  })

  it('writeCsv writes what Excel saved, read in its Mac code page, but where the platforms part', async () => {
    const { files } = await built()
    const known = KNOWN_WRITE[set] ?? {}
    const read = KNOWN_READ[set] ?? {}
    const decoder = new TextDecoder(tag === 'ru-RU' ? 'x-mac-cyrillic' : 'macintosh')
    const parted: string[] = []
    let compared = 0
    for (const c of stored.cases) {
      if (read[c.id]) continue
      compared++
      const excel = decoder.decode(Buffer.from(stored.answers[c.id].csv, 'base64'))
      if (engineAnswer(c, files.get(c.id)!, tag, now).csv !== excel) parted.push(c.id)
    }
    expect(parted.sort(), JSON.stringify(known)).toEqual(Object.keys(known).sort())
    expect(compared).toBeGreaterThan(parted.length)
  })
})

it('the sets of Excel for Mac are stored: files, held, files-ru, held-ru', () => {
  expect(sets.map((s) => `${s.tag} ${s.set}`)).toEqual(['en-US files', 'en-US held', 'ru-RU files-ru', 'ru-RU held-ru'])
})
