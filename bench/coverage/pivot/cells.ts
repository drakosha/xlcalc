/**
 * The set `cells`: what each cell of a pivot table is (`Range.PivotCell`: its type, field, item, data field and items
 * on the rows and columns), for `pivotCells`, over the books Excel saved for other sets, each opened as it is: the
 * pivot tables of `areas` in every form (without column fields, without row fields, value fields alone, none, no field,
 * report filters wrapped and across, the classic layout), of `basic` and of the three `corpus-shapes` sets.
 */
import { savedBooks } from './areas.js'
import type { PivotCase } from './cases.js'

/** The books of the sets as cases that read what each cell of their pivot tables is; `only` keeps those of the ids it passes. */
export const kindsOf = (sets: string[], only: (id: string) => boolean = () => true): PivotCase[] => savedBooks(sets, '').filter((c) => only(c.id)).map((c) => ({ ...c, kinds: true }))

export function buildPivotCases(): PivotCase[] {
  return kindsOf(['areas', 'basic', 'corpus-shapes', 'corpus-shapes-2', 'corpus-shapes-3'], (id) => !id.startsWith('areas-saved-'))
}
