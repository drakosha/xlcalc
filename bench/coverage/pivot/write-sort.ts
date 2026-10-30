/** The set `write-sort`: the books of `sort` written again by `writeXlsx`, as `write-1` writes those of `basic`; held out until the writer was done. */
import { buildPivotCases as sort } from './sort.js'
import { oursCases } from './ours.js'

export const buildPivotCases = () => oursCases('sort', 'xlsx', sort())
