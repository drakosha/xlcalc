/** The set `write-layout`: the books of `layout` written again by `writeXlsx`, as `write-1` writes those of `basic`; held out until the writer was done. */
import { buildPivotCases as layout } from './layout.js'
import { oursCases } from './ours.js'

export const buildPivotCases = () => oursCases('layout', 'xlsx', layout())
