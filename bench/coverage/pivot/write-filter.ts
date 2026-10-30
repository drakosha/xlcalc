/** The set `write-filter`: the books of `filter` written again by `writeXlsx`, as `write-1` writes those of `basic`; held out until the writer was done. */
import { buildPivotCases as filter } from './filter.js'
import { oursCases } from './ours.js'

export const buildPivotCases = () => oursCases('filter', 'xlsx', filter())
