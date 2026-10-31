/** The set `refresh-edit`: Excel's own books of `edit` opened, read, refreshed and read again, the control of `write-kept-edit` after its refresh. */
import { buildPivotCases as edit } from './edit.js'
import { oursCases } from './ours.js'

export const buildPivotCases = () => oursCases('edit', undefined, edit().filter((c) => c.ops.some((o) => 'edit' in o)))
