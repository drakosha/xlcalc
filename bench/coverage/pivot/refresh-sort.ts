/** The set `refresh-sort`: Excel's own books of `sort` opened, read, refreshed and read again, what the engine's files of `write-*` are compared with after their refresh. */
import { buildPivotCases as sort } from './sort.js'
import { oursCases } from './ours.js'

export const buildPivotCases = () => oursCases('sort', undefined, sort())
