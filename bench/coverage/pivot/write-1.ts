/**
 * The set `write-1`: each book of `basic` read and written again by the engine (`writeXlsx`, `ours.ts basic xlsx`
 * first), opened by Excel and read at once (no refresh on opening), then refreshed (`RefreshTable`) and read again.
 * Each read is to equal the area Excel showed for the book it saved (`test/pivot-write-excel.test.ts`).
 */
import { buildPivotCases as basic } from './basic.js'
import { oursCases } from './ours.js'

export const buildPivotCases = () => oursCases('basic', 'xlsx', basic())
