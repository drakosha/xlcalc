/** Statistical and compatibility functions, assembled from `./statistical/`. */
import type { FnDef } from '../args.js'
import { functions as compatibility } from './statistical/compatibility.js'
import { functions as conditional } from './statistical/conditional.js'
import { functions as descriptive } from './statistical/descriptive.js'
import { functions as distributions } from './statistical/distributions.js'
import { functions as inverse } from './statistical/inverse.js'
import { functions as regression } from './statistical/regression.js'
import { functions as tests } from './statistical/tests.js'

/**
 * Computed, but Excel's answer differs on a few extreme cases where Excel itself fails (#NUM! where the
 * quantile exists, digits lost at 10^9 degrees of freedom): outside the whitelist until those cases are
 * listed as known differences.
 */
const PENDING = new Set(['BINOM.INV', 'CHIINV', 'CHISQ.INV', 'CHISQ.INV.RT', 'F.INV', 'F.INV.RT', 'FINV', 'GAMMA.INV', 'GAMMAINV', 'T.INV', 'T.INV.2T', 'TINV'])

const all: Record<string, FnDef> = { ...descriptive, ...conditional, ...distributions, ...inverse, ...tests, ...regression, ...compatibility }

export const functions: Record<string, FnDef> = Object.fromEntries(Object.entries(all).filter(([name]) => !PENDING.has(name)))
