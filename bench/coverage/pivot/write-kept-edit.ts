/**
 * The set `write-kept-edit`: each case of `edit` from the book of its first phase, the cell edited in a session with
 * `xlcalc/pivot` and the book written by `writeKept` (`ours.ts edit kept-edit` first), opened by Excel, read at once,
 * refreshed and read again: as Excel showed its own book after the edit, and as that book refreshed (`refresh-edit`).
 */
import { buildPivotCases as edit } from './edit.js'
import { oursCases } from './ours.js'

export const buildPivotCases = () => oursCases('edit', 'kept-edit', edit())
