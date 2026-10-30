/** The group `lambda` of `xlcalc/functions` (`docs/function-groups.md`). */
import { functions } from '../functions/lambda.js'
import { functions as groupby } from '../functions/lookup/groupby.js'
import { group } from './group.js'

/** The group lambda: the functions that call a function for each value (MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY), GROUPBY and PIVOTBY. */
export const lambda = /* @__PURE__ */ group('lambda', functions, groupby)
