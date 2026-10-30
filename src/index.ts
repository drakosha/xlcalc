/**
 * A workbook formula engine (the package root). The contract (types only) is also available as the
 * `./contract` subpath; the implementation compiles formulas to closures (or to JS text) and computes over
 * the reference graph. The root has every function; `xlcalc/core` is the same API with the groups of `xlcalc/functions`
 * loaded as a host chooses (`docs/function-groups.md`).
 * The `./unstable/*` subpaths (internals, the synthetic book) are outside semver.
 */
import * as api from './api.js'
import { withAllFunctions } from './groups/all.js'

export * from './api.js'

// the calls that calculate load every group first: a bundle without them takes none of the groups
export const compute = /* @__PURE__ */ withAllFunctions(api.compute)
export const createSession = /* @__PURE__ */ withAllFunctions(api.createSession)
export const createInlineCalc = /* @__PURE__ */ withAllFunctions(api.createInlineCalc)
export const editStructure = /* @__PURE__ */ withAllFunctions(api.editStructure)
export const viewportOf = /* @__PURE__ */ withAllFunctions(api.viewportOf)
export const debugGeneratedCode = /* @__PURE__ */ withAllFunctions(api.debugGeneratedCode)
