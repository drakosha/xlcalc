/** The set `refresh-filter`: Excel's own books of `filter` opened, read, refreshed and read again, what the engine's files of `write-*` are compared with after their refresh. */
import { buildPivotCases as filter } from './filter.js'
import { oursCases } from './ours.js'

export const buildPivotCases = () => oursCases('filter', undefined, filter())
