/**
 * The subpath `xlcalc/functions`: the groups of rarely used functions, which `registerFunctions` of `xlcalc/core` loads
 * (`docs/function-groups.md`); the package root has them all. A bundle keeps the groups it imports.
 */
import type { FunctionGroup } from '../contract.js'
import { engineering } from './engineering.js'
import { finance } from './finance.js'
import { lambda } from './lambda.js'
import { statistics } from './statistics.js'
import { xml } from './xml.js'

export type { FunctionGroup, FunctionGroupName } from '../contract.js'
export { engineering, finance, lambda, statistics, xml }

/** Every group: `registerFunctions(...allFunctions)` loads all the functions the package root has. */
export const allFunctions: readonly FunctionGroup[] = [engineering, finance, lambda, statistics, xml]
