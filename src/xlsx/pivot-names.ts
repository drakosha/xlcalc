/**
 * The names a PivotTable's part gives what the model names otherwise, which its reader takes the other way: a
 * function's subtotal (`countA` for `count`), Show Values As, an area's axis.
 */
import type { PivotArea, PivotFunction, PivotShowAs } from '../contract.js'

/** The attribute of a reference's subtotal by function (`sumSubtotal`), and the item type of a subtotal (`t="sum"`). */
export const SUBTOTAL_ATTR: Record<PivotFunction, string> = { sum: 'sum', count: 'countA', average: 'avg', max: 'max', min: 'min', product: 'product', countNums: 'count', stdDev: 'stdDev', stdDevp: 'stdDevP', var: 'var', varp: 'varP' }

/** Show Values As where the file names it otherwise. */
export const SHOW_NAME: Partial<Record<PivotShowAs, string>> = { percentOfColumn: 'percentOfCol', percentOfParentColumn: 'percentOfParentCol' }

export const AXIS: Record<NonNullable<PivotArea['axis']>, string> = { rows: 'axisRow', columns: 'axisCol', pages: 'axisPage', values: 'axisValues' }

/** A table of names the other way: the model's by the file's. */
export const inverse = <K extends string>(o: Partial<Record<K, string>>): Record<string, K> => Object.fromEntries(Object.entries(o).map(([k, v]) => [v, k]))
