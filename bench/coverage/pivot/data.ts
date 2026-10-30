/**
 * The source data of the pivot sets: deterministic rows on the sheet `D` of each case's book, with the kinds of values a
 * pivot cache has to sort out (text, numbers, dates and times, empty cells, errors, one item in two cases, numbers kept
 * as text, booleans, long text). The book also has the sheet `P`, where the pivot tables go, and `G`, for formulas
 * that read them. The same rows can be a range, a table (`Sales`) or a defined name (`SalesData`) to the cache.
 */
import type { BookInput, CellInput, SheetInput } from '../../../src/contract.js'
import { formatAddress } from '../../../src/model.js'

/** Where a cache takes its rows: the range itself, the table over it, or the defined name over it. */
export type SourceKind = 'range' | 'table' | 'name'

/** A source column: its header and its value in each row (undefined for an empty cell). */
export interface SourceColumn {
  name: string
  values: (CellInput | undefined)[]
}

/** A date as Excel's serial number (the 1900 system). */
export const serial = (y: number, m: number, d: number, h = 0, min = 0): number => Date.UTC(y, m - 1, d, h, min) / 86_400_000 + 25_569

/** A pseudo-random sequence from a seed (a linear congruential generator), the same on every run. */
function sequence(seed: number): () => number {
  let x = seed >>> 0
  return () => {
    x = (Math.imul(x, 1_664_525) + 1_013_904_223) >>> 0
    return x / 2 ** 32
  }
}

const LONG = 'Dragonfruit from the far southern orchards, packed in wooden crates of twelve and shipped by sea'
const DATE = 'yyyy-mm-dd'
const WHEN = 'yyyy-mm-dd hh:mm'

/**
 * The rows of `sales`: Region (four regions, `north` beside `North`, one empty), Product (with a long name and a number
 * kept as text), Date (dates over 2024 to 2025, one empty), Amount (numbers with empty cells, text, a number kept as
 * text, a boolean, #N/A and #DIV/0!), Price (numbers with empty cells and text, no errors), Qty (whole numbers), Code
 * (numbers, one kept as text), Flag (booleans), Note (long and short text, empty cells), When (dates with times).
 */
export function salesColumns(rows = 48): SourceColumn[] {
  const rnd = sequence(20_261_030)
  const pick = <T>(xs: readonly T[]) => xs[Math.floor(rnd() * xs.length)]
  const regions = ['North', 'South', 'East', 'West'] as const
  const products = ['Apple', 'Banana', 'Cherry', LONG, '100'] as const
  const cols: Record<string, (CellInput | undefined)[]> = { Region: [], Product: [], Date: [], Amount: [], Price: [], Qty: [], Code: [], Flag: [], Note: [], When: [] }
  for (let k = 0; k < rows; k++) {
    const region = k === 9 ? undefined : k % 10 === 4 ? 'north' : regions[k % 4]
    cols.Region.push(region === undefined ? undefined : { value: region })
    cols.Product.push({ value: k % 9 === 8 ? LONG : pick(products.slice(0, 3)) })
    if (k === 31) cols.Product[k] = { value: '100' }
    cols.Date.push(k === 17 ? undefined : { value: serial(2024 + (k % 2), 1 + ((k * 5) % 12), 1 + ((k * 7) % 28)), numberFormat: DATE })
    const amount = Math.round((rnd() * 900 - 100) * 100) / 100
    cols.Amount.push(
      k % 11 === 3 ? undefined
      : k === 5 ? { value: 'n/a' }
      : k === 7 ? { value: '12' }
      : k === 30 ? { value: true }
      : k === 20 ? { formula: 'NA()' }
      : k === 41 ? { formula: '1/0' }
      : { value: amount },
    )
    cols.Price.push(k % 7 === 2 ? undefined : k === 12 ? { value: 'free' } : { value: Math.round(rnd() * 5000) / 100 })
    cols.Qty.push({ value: 1 + Math.floor(rnd() * 9) })
    cols.Code.push({ value: k === 22 ? '20' : [10, 20, 30][k % 3] })
    cols.Flag.push({ value: k % 3 !== 1 })
    cols.Note.push(k % 4 === 1 ? undefined : { value: k % 8 === 0 ? `${LONG}, row ${k + 1}` : pick(['ok', 'late', 'OK', 'check']) })
    cols.When.push({ value: serial(2025, 1 + (k % 3), 1 + (k % 5), (k * 5) % 24, (k * 13) % 60), numberFormat: WHEN })
  }
  return Object.entries(cols).map(([name, values]) => ({ name, values }))
}

/** The data sets by name. */
export const DATA: Record<string, () => SourceColumn[]> = { sales: () => salesColumns() }

/** The range the columns take on `D`, from A1 with the header row (`A1:J49`). */
export const sourceRange = (cols: SourceColumn[]): string => `A1:${formatAddress(cols[0].values.length, cols.length - 1)}`

/** A source as the cache's `SourceData` names it: the range in R1C1 with its sheet, the table's name or the defined name. */
export function sourceText(kind: SourceKind, cols: SourceColumn[]): string {
  if (kind === 'table') return 'Sales'
  if (kind === 'name') return 'SalesData'
  return `D!R1C1:R${cols[0].values.length + 1}C${cols.length}`
}

/** The book a case starts from: the columns on `D` (in a table or under a defined name when the source is one), empty `P` and `G`. */
export function sourceBook(cols: SourceColumn[], kind: SourceKind): BookInput {
  const cells: Record<string, CellInput> = {}
  cols.forEach((c, j) => {
    cells[formatAddress(0, j)] = { value: c.name }
    c.values.forEach((v, i) => v && (cells[formatAddress(i + 1, j)] = v))
  })
  const range = sourceRange(cols)
  const data: SheetInput = { name: 'D', cells, ...(kind === 'table' ? { tables: [{ name: 'Sales', range, columns: cols.map((c) => ({ name: c.name })) }] } : {}) }
  return {
    sheets: [data, { name: 'P', cells: {} }, { name: 'G', cells: {} }],
    ...(kind === 'name' ? { names: [{ name: 'SalesData', ref: `D!${range.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}` }] } : {}),
  }
}
