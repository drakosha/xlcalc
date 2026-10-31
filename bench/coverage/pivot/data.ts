/**
 * The source data of the pivot sets: deterministic rows on the sheet `D` of each case's book, with the kinds of values a
 * pivot cache has to sort out (text, numbers, dates and times, empty cells, errors, one item in two cases, numbers kept
 * as text, booleans, long text). The book also has the sheet `P`, where the pivot tables go, and `G`, for formulas
 * that read them. The same rows can be a range, a table (`Sales`) or a defined name (`SalesData`) to the cache.
 */
import type { BookInput, CellInput, SheetInput } from '../../../src/contract.js'
import { formatAddress } from '../../../src/model.js'
import { ORDER_DATA } from './order-data.js'
import { numericOrderColumns } from './numeric-order-data.js'
import { filterOrderColumns } from './filter-order-data.js'
import { kindColumns, kindColumns2 } from './kinds-data.js'

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

/** What tells two data sets of the same columns apart: rows, seed, items and where the odd cells fall. */
interface Variant {
  rows: number
  seed: number
  regions: readonly string[]
  /** The region written in lower case every tenth row from the fifth, and the row with no region. */
  lower: string
  noRegion: number
  products: readonly string[]
  long: string
  /** The product kept as a number in text, and its row. */
  numberText: [string, number]
  codes: readonly number[]
  /** The row whose code is text, and the rows of the odd Amounts: text, a number as text, a boolean, #N/A, #DIV/0!. */
  codeText: number
  odd: [text: number, numberText: number, bool: number, na: number, div0: number]
  noDate: number
}

const SALES: Variant = { rows: 48, seed: 20_261_030, regions: ['North', 'South', 'East', 'West'], lower: 'north', noRegion: 9, products: ['Apple', 'Banana', 'Cherry'], long: LONG, numberText: ['100', 31], codes: [10, 20, 30], codeText: 22, odd: [5, 7, 30, 20, 41], noDate: 17 }

/** The rows the held-out sets start from: the columns of `sales`, other rows, items and odd cells. */
const STOCK: Variant = { rows: 60, seed: 20_261_031, regions: ['Central', 'West', 'east', 'South', 'North'], lower: 'west', noRegion: 23, products: ['Plum', 'Pear', 'Apple', 'Fig'], long: 'Quince preserved in a syrup of honey and lemon peel, sold by the jar at the market on Saturdays', numberText: ['7', 44], codes: [5, 15, 25, 35], codeText: 13, odd: [11, 2, 38, 50, 27], noDate: 40 }

/**
 * The rows of a variant: Region (regions, one also in lower case, one row empty), Product (with a long name and a number
 * kept as text), Date (dates over 2024 to 2025, one empty), Amount (numbers with empty cells, text, a number kept as
 * text, a boolean, #N/A and #DIV/0!), Price (numbers with empty cells and text, no errors), Qty (whole numbers), Code
 * (numbers, one kept as text), Flag (booleans), Note (long and short text, empty cells), When (dates with times).
 */
function variantColumns(v: Variant): SourceColumn[] {
  const rnd = sequence(v.seed)
  const pick = <T>(xs: readonly T[]) => xs[Math.floor(rnd() * xs.length)]
  const cols: Record<string, (CellInput | undefined)[]> = { Region: [], Product: [], Date: [], Amount: [], Price: [], Qty: [], Code: [], Flag: [], Note: [], When: [] }
  const [text, numberText, bool, na, div0] = v.odd
  for (let k = 0; k < v.rows; k++) {
    const region = k === v.noRegion ? undefined : k % 10 === 4 ? v.lower : v.regions[k % v.regions.length]
    cols.Region.push(region === undefined ? undefined : { value: region })
    // the pick is drawn on the row of the number kept as text too, so the sequence is the same either way
    const product = k % 9 === 8 ? v.long : pick(v.products)
    cols.Product.push({ value: k === v.numberText[1] ? v.numberText[0] : product })
    cols.Date.push(k === v.noDate ? undefined : { value: serial(2024 + (k % 2), 1 + ((k * 5) % 12), 1 + ((k * 7) % 28)), numberFormat: DATE })
    const amount = Math.round((rnd() * 900 - 100) * 100) / 100
    cols.Amount.push(
      k % 11 === 3 ? undefined
      : k === text ? { value: 'n/a' }
      : k === numberText ? { value: '12' }
      : k === bool ? { value: true }
      : k === na ? { formula: 'NA()' }
      : k === div0 ? { formula: '1/0' }
      : { value: amount },
    )
    cols.Price.push(k % 7 === 2 ? undefined : k === 12 ? { value: 'free' } : { value: Math.round(rnd() * 5000) / 100 })
    cols.Qty.push({ value: 1 + Math.floor(rnd() * 9) })
    cols.Code.push({ value: k === v.codeText ? String(v.codes[1]) : v.codes[k % v.codes.length] })
    cols.Flag.push({ value: k % 3 !== 1 })
    cols.Note.push(k % 4 === 1 ? undefined : { value: k % 8 === 0 ? `${v.long}, row ${k + 1}` : pick(['ok', 'late', 'OK', 'check']) })
    cols.When.push({ value: serial(2025, 1 + (k % 3), 1 + (k % 5), (k * 5) % 24, (k * 13) % 60), numberFormat: WHEN })
  }
  return Object.entries(cols).map(([name, values]) => ({ name, values }))
}

export const salesColumns = (): SourceColumn[] => variantColumns(SALES)

/**
 * The rows of `shapes`, few and made by hand for the shapes of the PivotTables a corpus showed: Region first comes as B,
 * D, A, C (its order sorted A, B, C, D, E, F), and A, B and D sum 10 of Qty each, 4 of it with X and 6 with Y, so a sort
 * by a value has ties over the products and within one; E holds no Qty, F only the product Z; Code numbers.
 */
export function shapeColumns(): SourceColumn[] {
  const rows: [string, string, number, number | undefined, number][] = [
    ['B', 'X', 5, 4, 1.5],
    ['D', 'X', 15, 4, 2.25],
    ['A', 'Y', 5, 6, 3],
    ['C', 'X', 25, 5, 4.75],
    ['E', 'Z', 35, undefined, 5],
    ['B', 'Y', 5, 6, 6.5],
    ['D', 'Y', 15, 6, 7],
    ['A', 'X', 25, 4, 8.25],
    ['F', 'Z', 25, 3, 9],
  ]
  const names = ['Region', 'Product', 'Code', 'Qty', 'Price']
  return names.map((name, j) => ({ name, values: rows.map((r) => (r[j] === undefined ? undefined : { value: r[j] as string | number })) }))
}

/**
 * The rows of `blanks`, made by hand: Product of text, a number and empty cells, B and D with an empty Product alone, B
 * with Flag FALSE alone; for an item whose records all lie in a hidden item of another field.
 */
export function blankColumns(): SourceColumn[] {
  const rows: [string, string | number | undefined, number, number, boolean][] = [
    ['A', 'X', 1, 2, true],
    ['A', undefined, 2, 3, true],
    ['B', undefined, 2, 4, false],
    ['C', 10, 1, 5, true],
    ['C', 'X', 3, 1, false],
    ['D', undefined, 4, 6, true],
  ]
  const names = ['Region', 'Product', 'Code', 'Qty', 'Flag']
  return names.map((name, j) => ({ name, values: rows.map((r) => (r[j] === undefined ? undefined : { value: r[j] as string | number | boolean })) }))
}

/**
 * The rows of `blanks2`, made by hand: Region with empty cells whose Product is empty too, Product of numbers, text and
 * empty cells, Flag FALSE on some rows of A and of the empty Region.
 */
export function blankColumns2(): SourceColumn[] {
  const rows: [string | undefined, string | number | undefined, number, number, boolean][] = [
    ['A', 1, 5, 2, true],
    ['A', 'X', 6, 3, false],
    [undefined, undefined, 5, 4, true],
    [undefined, undefined, 6, 7, false],
    ['B', 2, 7, 5, false],
    ['B', undefined, 7, 1, true],
    ['A', 'X', 5, 8, true],
  ]
  const names = ['Region', 'Product', 'Code', 'Qty', 'Flag']
  return names.map((name, j) => ({ name, values: rows.map((r) => (r[j] === undefined ? undefined : { value: r[j] as string | number | boolean })) }))
}

/** The data sets by name. */
export const DATA: Record<string, () => SourceColumn[]> = { ...ORDER_DATA, 'filter-numeric': () => filterOrderColumns(), 'filter-errors': () => filterOrderColumns(true), 'numeric-order': numericOrderColumns, 'numeric-held': () => numericOrderColumns(0xbb67ae85), sales: salesColumns, stock: () => variantColumns(STOCK), shapes: shapeColumns, blanks: blankColumns, blanks2: blankColumns2, kinds: kindColumns, kinds2: kindColumns2 }

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
