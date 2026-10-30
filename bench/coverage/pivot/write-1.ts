/**
 * The set `write-1`: each book of `basic` read and written again by the engine (`writeXlsx`), opened by Excel and read
 * at once (no refresh on opening), then refreshed (`RefreshTable`) and read again. Each read is to equal the area Excel
 * showed for the book it saved (`test/pivot-write-excel.test.ts`).
 */
import type { PivotCase } from './cases.js'
import { buildPivotCases as basic } from './basic.js'

export function buildPivotCases(): PivotCase[] {
  return basic().map((c): PivotCase => ({ id: c.id, from: `basic/${c.id}`, rewrite: 'xlsx', ops: [{ read: true }, { refresh: true }] }))
}
