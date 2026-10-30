/**
 * The subpath `xlcalc/core`: the package root's API without the groups of rarely used functions, which a host loads from
 * `xlcalc/functions` by `registerFunctions` (`docs/function-groups.md`). A formula that calls a function of a group not
 * loaded is `#NAME?` with the check `function_not_loaded`; the checks of formula text know every function.
 */
import * as api from './api.js'
import { GROUP_ARITY } from './groups/manifest.js'
import * as registry from './groups/registry.js'
import { before } from './strict-args.js'

export * from './api.js'

// the groups not loaded are known by their functions' arguments, and a pool's threads load the groups registered, once
// the host calculates or registers through this entry (a bundle drops a call at its top with the calls it does not use)
const known = () => registry.useManifest(GROUP_ARITY)
const light = <F extends (...args: never[]) => unknown>(f: F): F => before(f, known)

export const compute = /* @__PURE__ */ light(api.compute)
export const createSession = /* @__PURE__ */ light(api.createSession)
export const createInlineCalc = /* @__PURE__ */ light(api.createInlineCalc)
export const editStructure = /* @__PURE__ */ light(api.editStructure)
export const viewportOf = /* @__PURE__ */ light(api.viewportOf)
export const debugGeneratedCode = /* @__PURE__ */ light(api.debugGeneratedCode)

/**
 * Loads groups of `xlcalc/functions` for the calculations from now on, a pool's threads too:
 * `registerFunctions(engineering, xml)`, or `registerFunctions(...allFunctions)` for all of them. A session opened before
 * builds its workbook again at its next call, calculating the formulas over them.
 */
export const registerFunctions = /* @__PURE__ */ light(registry.registerFunctions)
