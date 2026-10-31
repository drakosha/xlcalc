/** The set `write-slicer`: the books of `slicer` (slicers and timelines) written again by `writeXlsx`, as `write-1` writes those of `basic`. */
import { buildPivotCases as slicer } from './slicer.js'
import { oursCases } from './ours.js'

export const buildPivotCases = () => oursCases('slicer', 'xlsx', slicer())
