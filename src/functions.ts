/**
 * The function whitelist, assembled from the category files in `src/functions/`. A function gets into its
 * category's registry only when all of its cases in the Excel acceptance test (`test/excel.test.ts`) match
 * Excel en-US, or with its own semantics tests where the measurement has no cases.
 */
import { specOf, type FnDef, type FunctionSpec } from './args.js'
import { functions as database } from './functions/database.js'
import { functions as date } from './functions/date.js'
import { functions as engineering } from './functions/engineering.js'
import { functions as financial } from './functions/financial.js'
import { functions as information } from './functions/information.js'
import { functions as logical } from './functions/logical.js'
import { functions as lookup } from './functions/lookup.js'
import { functions as math } from './functions/math.js'
import { functions as statistical } from './functions/statistical.js'
import { functions as text } from './functions/text.js'

export type { FunctionSpec } from './args.js'

const CATEGORIES = { database, date, engineering, financial, information, logical, lookup, math, statistical, text }

function assemble(): Record<string, FnDef> {
  const out: Record<string, FnDef> = {}
  for (const [category, defs] of Object.entries(CATEGORIES))
    for (const [name, def] of Object.entries(defs)) {
      if (out[name]) throw new Error(`function ${name} is defined twice (again in ${category})`)
      out[name] = def
    }
  return out
}

/** Implementations and argument kinds by function name. */
export const DEFS: Readonly<Record<string, FnDef>> = assemble()

/** Argument counts and range positions by function name, for the parser and the code generator. */
export const FUNCTIONS: Readonly<Record<string, FunctionSpec>> = Object.fromEntries(Object.entries(DEFS).map(([name, def]) => [name, specOf(name, def)]))

export const FUNCTION_NAMES: readonly string[] = Object.keys(FUNCTIONS).sort()
