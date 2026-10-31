/**
 * The set `over-data`: a PivotTable made over cells that are not empty (`PT2` at `P!H3`, over the cache of `PT1`), as
 * Excel makes it with its alerts off, which answers its question to replace the destination's contents: a value, a
 * formula, a number format alone and with a value, merged cells and a table in the place of the new PivotTable, a
 * value in its top-left cell; then its fields placed, and a value below that it grows over.
 */
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const made: PivotOp[] = [{ create: 'PT2', at: 'P!H3' }, { read: true }, { rows: ['Region'], pt: 'PT2' }, { value: { field: 'Qty' }, pt: 'PT2' }]
  const c = (id: string, before: PivotOp[], after: PivotOp[] = made): PivotCase => ({ id, ops: [...before, ...after] })
  return [
    c('value', [{ edit: 'P!I5', to: 'x' }]),
    c('formula', [{ edit: 'P!H10', to: { formula: '1+1' } }]),
    c('format', [{ cellFormat: 'P!I6', numberFormat: '0.000' }]),
    c('value-format', [{ edit: 'P!I6', to: 5 }, { cellFormat: 'P!I6', numberFormat: '0.000' }]),
    c('merged', [{ edit: 'P!H6', to: 'm' }, { merge: 'P!H6:I6' }]),
    c('table', [{ edit: 'P!I8', to: 'A' }, { edit: 'P!J8', to: 'B' }, { edit: 'P!I9', to: 1 }, { edit: 'P!J9', to: 2 }, { addTable: 'P!I8:J9', name: 'T1' }]),
    c('top-left', [{ edit: 'P!H3', to: 'x' }]),
    c('grows-over', [{ edit: 'P!H25', to: 'x' }], [...made, { rows: ['Product'], pt: 'PT2' }]),
  ]
}
