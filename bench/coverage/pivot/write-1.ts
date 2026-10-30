/**
 * The set `write-1`: each book of `basic` read and written again by the engine (`writeXlsx`, `ours.ts basic xlsx`
 * first), opened by Excel and read at once (no refresh on opening), then refreshed (`RefreshTable`) and read again.
 * Each read is to equal the area Excel showed for the book it saved (`test/pivot-write-excel.test.ts`).
 */
import type { PivotCase } from './cases.js'
import { buildPivotCases as basic } from './basic.js'
import { oursPath } from './ours.js'

export function buildPivotCases(): PivotCase[] {
  return basic().map((c): PivotCase => ({ id: c.id, file: oursPath('xlsx', 'basic', c.id), ops: [{ read: true }, { refresh: true }] }))
}
