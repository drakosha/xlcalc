/**
 * Database functions (DSUM, DCOUNT, …). None is in the whitelist yet: formula.js reads the database by
 * columns with the field number shifted by one, so they are to be written here on the engine's criteria
 * (`src/criteria.ts`) and verified against the Excel cases.
 */
import type { FnDef } from '../args.js'

export const functions: Record<string, FnDef> = {}
