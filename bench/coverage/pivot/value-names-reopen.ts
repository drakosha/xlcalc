/**
 * The set `value-names-reopen`, run in en-US and ru-RU after `value-names`: the books `value-names` saved after naming
 * their value field, opened again, and the function changed to Count, so that a name Excel keeps no mark of in the file
 * is told from its default name by its text alone or not.
 */
import type { PivotCase } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  return ['default', 'own', 'other-default-en', 'sum-text-en', 'sum-text-ru', 'named-back', 'lower-case'].map((id) => ({ id, from: `value-names/${id}-1`, ops: [{ setValue: 1, fn: 'count' }] }))
}
