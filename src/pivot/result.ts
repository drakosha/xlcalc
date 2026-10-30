/**
 * What a PivotTable's calculation gives its layout: the items shown on each axis as a tree in the order shown, the
 * value of any cell of the cube by the items of the rows and of the columns, and the items shown on the Filters area.
 * Hidden and filtered items are already out of the trees, and sorting is applied; Show Values As calculates over
 * `value`, which answers for any items, shown or not.
 */
import type { CellValue, PivotAxisField, PivotCache, PivotFunction, PivotTable } from '../contract.js'

/** An item shown on an axis, with the items of the next field under it. */
export interface PivotNode {
  /** The cache field's index, or `values` for a value field when the value fields lie across this axis. */
  field: PivotAxisField
  /** The index into the PivotTable field's `items`, or into `PivotTable.values` for `values`. */
  item: number
  /** The next field's items under this one, in the order shown; empty on the axis's last field. */
  children: PivotNode[]
  /** No record reaches it: it is shown only because its field shows items with no data (`showEmpty`). */
  empty?: boolean
}

/** An axis (rows or columns): its fields in order and its first field's items as shown. */
export interface PivotAxisResult {
  /** The axis's fields, `values` among them where the value fields lie across it. */
  fields: PivotAxisField[]
  /** The first field's items in the order shown; empty for an axis with no fields. */
  items: PivotNode[]
}

/** A field on the Filters area and the items it lets through. */
export interface PivotPageResult {
  field: number
  /** Indexes into the PivotTable field's `items` of the items let through, in the field's order. */
  items: number[]
  /** Every item is let through (the field shows `(All)`). */
  all: boolean
}

/**
 * A PivotTable calculated. A path is the items of an axis's fields other than `values`, from the first, each an index
 * into the PivotTable field's `items`: the whole path is a line of data, a shorter one the subtotal of its last item,
 * the empty path the grand total.
 */
export interface PivotResult {
  /** The PivotTable with its fields' items as the cache lists them now (new items in, items gone out). */
  table: PivotTable
  /** The cache it was calculated from, refreshed when the calculation refreshed it. */
  cache: PivotCache
  rows: PivotAxisResult
  columns: PivotAxisResult
  pages: PivotPageResult[]
  /**
   * The value of the value field `value` (its index in `PivotTable.values`) at the paths of the rows and of the
   * columns, as Excel shows it before Show Values As: a number, an error, or null where no record lies. `fn` is a
   * subtotal's own function in place of the value field's, for a field with several subtotals.
   */
  value(rows: readonly number[], columns: readonly number[], value: number, fn?: PivotFunction): CellValue
  /** Some record that the filters let through lies at the paths. */
  hasData(rows: readonly number[], columns: readonly number[]): boolean
}

/** What a calculation takes besides the PivotTable and its cache. */
export interface PivotComputeOptions {
  /** Today as a serial number, the session's clock (as `TODAY`), for the date filters relative to it. */
  today?: number
  /** The workbook counts dates from 1904. */
  date1904?: boolean
}
