/** The set `refresh-slicer`: Excel's own books of `slicer` opened, read, refreshed and read again, the control of `write-slicer` after its refresh. */
import { buildPivotCases as slicer } from './slicer.js'
import { oursCases } from './ours.js'

export const buildPivotCases = () => oursCases('slicer', undefined, slicer())
