/** Statistical and compatibility functions of the core, assembled from `./statistical/`. */
import type { FnDef } from '../args.js'
import { functions as compatibility } from './statistical/compatibility.js'
import { functions as conditional } from './statistical/conditional.js'
import { functions as descriptive } from './statistical/descriptive.js'
import { functions as regression } from './statistical/regression.js'

// the distributions, their inverses, the tests and the fits are the group `statistics` (`groups/statistics.ts`)
export const functions: Record<string, FnDef> = { ...descriptive, ...conditional, ...regression, ...compatibility }
