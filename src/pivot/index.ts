/**
 * `xlcalc/pivot`: the PivotTable engine, which `registerPivot` loads for every session, a pool's threads too: refresh of
 * pivot caches from their sources, the value fields summarized, filters, sorting, and the layout of PivotTables on their
 * sheets (`docs/pivot-tables.md`).
 */
import type { PivotEngine } from '../contract.js'
import type { PivotEngineParts } from '../pivot-level.js'
import { checkPivotShapes } from './check.js'
import { editPivotCache } from './cache-ops.js'
import { refreshPivots } from './engine.js'
import { pivotEntry } from './entry.js'
import { valuesTaken } from '../structure/pivot-values.js'

/** The PivotTable engine, which `registerPivot(pivotEngine)` loads: caches refreshed from their sources, PivotTables calculated and laid out ([PivotTables](pivot-tables.md)). */
export const pivotEngine: PivotEngine = { kind: 'xlcalc-pivot', refresh: refreshPivots, check: checkPivotShapes, entry: pivotEntry, edit: editPivotCache, taken: valuesTaken } as PivotEngineParts
