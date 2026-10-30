/** A PivotTable's calculation over its cache: the axes' items shown and the values of the cube (`PivotResult`). */
import type { PivotCache, PivotTable } from '../contract.js'
import type { PivotComputeOptions, PivotResult } from './result.js'

/** Calculates a PivotTable over its cache's records. */
export function computePivot(table: PivotTable, cache: PivotCache, options: PivotComputeOptions = {}): PivotResult {
  void options
  return {
    table,
    cache,
    rows: { fields: table.rows ?? [], items: [] },
    columns: { fields: table.columns ?? [], items: [] },
    pages: (table.pages ?? []).map((p) => ({ field: p.field, items: [], all: p.item === undefined })),
    value: () => null,
    hasData: () => false,
  }
}
