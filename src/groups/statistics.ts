/** The group `statistics` of `xlcalc/functions` (`docs/function-groups.md`). */
import { functions as distributions } from '../functions/statistical/distributions.js'
import { functions as ets } from '../functions/statistical/ets.js'
import { functions as fit } from '../functions/statistical/fit.js'
import { functions as inverse } from '../functions/statistical/inverse.js'
import { functions as old } from '../functions/statistical/old-distributions.js'
import { functions as tests } from '../functions/statistical/tests.js'
import { group } from './group.js'

/** The group statistics: the distributions, their inverses, the tests, LINEST and its kin, and their old names. */
export const statistics = /* @__PURE__ */ group('statistics', distributions, inverse, tests, fit, old, ets)
