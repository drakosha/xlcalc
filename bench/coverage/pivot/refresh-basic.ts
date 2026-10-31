/** The set `refresh-basic`: Excel's own books of `basic` opened, read, refreshed and read again, what the engine's files of `write-*` are compared with after their refresh. */
import { buildPivotCases as basic } from './basic.js'
import { oursCases } from './ours.js'

export const buildPivotCases = () => oursCases('basic', undefined, basic())
