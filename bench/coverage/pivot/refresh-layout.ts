/** The set `refresh-layout`: Excel's own books of `layout` opened, read, refreshed and read again, what the engine's files of `write-*` are compared with after their refresh. */
import { buildPivotCases as layout } from './layout.js'
import { oursCases } from './ours.js'

export const buildPivotCases = () => oursCases('layout', undefined, layout())
