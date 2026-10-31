/**
 * The set `areas`: the ranges of a pivot table's areas as Excel gives them (`TableRange1`, `TableRange2`,
 * `DataBodyRange`, `PageRange`, `RowRange`, `ColumnRange`), for `PivotLocation.areas`. Pivot tables in each form:
 * compact, outline and tabular, with and without column fields, without row fields, values alone, value fields on the
 * rows, no value field, no field at all, report filters in one column, in several (`PageFieldWrapCount`) and across
 * first; then the books Excel saved for the sets `layout` and `corpus-shapes`, each opened as it is.
 */
import { readdirSync } from 'node:fs'
import type { PivotCase, PivotOp } from './cases.js'

/** The books Excel saved at the end of each case of the sets named (`<set>/<id>.xlsx`), each a case opened as it is. */
export function savedBooks(sets: string[], prefix: string): PivotCase[] {
  return sets.flatMap((set) => {
    const dir = `test/excel/win/pivot/${set}`
    return readdirSync(dir)
      .filter((f) => /^[^.]+\.xlsx$/.test(f) && !/-\d+\.xlsx$/.test(f))
      .sort()
      .map((f) => ({ id: `${prefix}${set}-${f.slice(0, -'.xlsx'.length)}`, file: `${dir}/${f}`, ops: [] }))
  })
}

export function buildPivotCases(): PivotCase[] {
  const qty: PivotOp = { value: { field: 'Qty' } }
  const price: PivotOp = { value: { field: 'Price' } }
  const c = (id: string, ops: PivotOp[]): PivotCase => ({ id, ops })
  return [
    c('compact', [{ rows: ['Region'] }, qty]),
    c('compact-two-rows', [{ rows: ['Region', 'Product'] }, qty]),
    c('compact-columns', [{ rows: ['Region'] }, { columns: ['Flag'] }, qty]),
    c('compact-two-columns', [{ rows: ['Region'] }, { columns: ['Flag', 'Code'] }, qty]),
    c('outline', [{ rows: ['Region', 'Product'] }, qty, { layout: 'outline' }]),
    c('outline-columns', [{ rows: ['Region', 'Product'] }, { columns: ['Flag'] }, qty, { layout: 'outline' }]),
    c('tabular', [{ rows: ['Region', 'Product'] }, qty, { layout: 'tabular' }]),
    c('tabular-columns', [{ rows: ['Region', 'Product'] }, { columns: ['Flag'] }, qty, price, { layout: 'tabular' }]),
    c('two-values', [{ rows: ['Region'] }, qty, price]),
    c('two-values-rows', [{ rows: ['Region'] }, qty, price, { valuesOn: 'row' }]),
    c('no-rows', [{ columns: ['Flag'] }, qty]),
    c('no-rows-two-values', [{ columns: ['Flag'] }, qty, price]),
    c('value-alone', [qty]),
    c('values-alone', [qty, price]),
    c('rows-alone', [{ rows: ['Region'] }]),
    c('rows-columns-no-value', [{ rows: ['Region'] }, { columns: ['Flag'] }]),
    c('columns-alone', [{ columns: ['Flag'] }]),
    c('empty', [{ refresh: true }]),
    c('page', [{ pages: ['Product'] }, { rows: ['Region'] }, qty]),
    c('pages-three', [{ pages: ['Product', 'Flag', 'Code'] }, { rows: ['Region'] }, qty]),
    c('pages-wrap', [{ pages: ['Product', 'Flag', 'Code'] }, { rows: ['Region'] }, qty, { table: { pagesWrap: 2 } }]),
    c('pages-across', [{ pages: ['Product', 'Flag', 'Code'] }, { rows: ['Region'] }, qty, { table: { pagesAcross: true } }]),
    c('pages-across-wrap', [{ pages: ['Product', 'Flag', 'Code'] }, { rows: ['Region'] }, qty, { table: { pagesAcross: true, pagesWrap: 2 } }]),
    c('page-alone', [{ pages: ['Product'] }]),
    c('page-value-alone', [{ pages: ['Product'] }, qty]),
    c('page-tabular-columns', [{ pages: ['Note'] }, { rows: ['Region', 'Product'] }, { columns: ['Flag'] }, qty, { layout: 'tabular' }]),
    c('no-grand', [{ rows: ['Region'] }, { columns: ['Flag'] }, qty, { table: { rowGrand: false, columnGrand: false } }]),
    c('classic', [{ table: { classic: true } }, { layout: 'tabular' }, { rows: ['Region'] }, qty]),
    ...savedBooks(['layout', 'corpus-shapes'], 'saved-'),
  ]
}
