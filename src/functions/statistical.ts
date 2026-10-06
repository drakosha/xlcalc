/** Statistical and compatibility functions, assembled from `./statistical/`. */
import type { FnDef } from '../args.js'
import { functions as compatibility } from './statistical/compatibility.js'
import { functions as conditional } from './statistical/conditional.js'
import { functions as descriptive } from './statistical/descriptive.js'
import { functions as distributions } from './statistical/distributions.js'
import { functions as inverse } from './statistical/inverse.js'
import { functions as regression } from './statistical/regression.js'
import { functions as tests } from './statistical/tests.js'

export const functions: Record<string, FnDef> = { ...descriptive, ...conditional, ...distributions, ...inverse, ...tests, ...regression, ...compatibility }
