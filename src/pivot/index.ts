/**
 * `xlcalc/pivot`: the PivotTable engine, which `registerPivot` loads for every session, a pool's threads too: refresh of
 * pivot caches from their sources, the value fields summarized, filters, sorting, and the layout of PivotTables on their
 * sheets (`docs/pivot-tables.md`).
 */
import type { PivotEngine } from '../contract.js'
import type { PivotEngineParts } from '../pivot-level.js'
import { refreshPivots } from './engine.js'

/** The engine: `registerPivot(pivotEngine)`. */
export const pivotEngine: PivotEngine = { kind: 'xlcalc-pivot', refresh: refreshPivots } as PivotEngineParts
