/**
 * GETPIVOTDATA over the books Excel saved for another set: a case opens one (`basic/columns`), types its formulas into
 * `G!A1` down (as Excel 365 types them), and some at `P!K1` down, beside the pivot table, and reads them back after
 * Excel computes them.
 */
import type { PivotCase, PivotOp } from './cases.js'

/** A case of formulas over a saved book: those on `G` from A1, those on `P` from K1, values `G!H1` down for references. */
export function gpdCase(id: string, from: string, onG: string[], onP: string[] = [], values: (number | string)[] = []): PivotCase {
  const ops: PivotOp[] = [
    ...values.map((v, i): PivotOp => ({ edit: `G!H${i + 1}`, to: v })),
    ...onG.map((f, i): PivotOp => ({ edit: `G!A${i + 1}`, to: { formula: f } })),
    ...onP.map((f, i): PivotOp => ({ edit: `P!K${i + 1}`, to: { formula: f } })),
  ]
  return { id, from, ops, watch: [`G!A1:A${onG.length}`, ...(onP.length ? [`P!K1:K${onP.length}`] : [])] }
}

/** `GETPIVOTDATA(data, P!$A$3, pairs...)` with the arguments as written. */
export const gpd = (data: string, ...pairs: string[]): string => `GETPIVOTDATA(${[data, 'P!$A$3', ...pairs].join(',')})`
