/** Every group loaded before a call of the package root that calculates (`index.ts`): the root has all the functions. */
import { before } from '../strict-args.js'
import { allFunctions } from './index.js'
import { loadGroups } from './registry.js'

let all = false

/** `f`, by its name and number of parameters, loading every group at its first call. */
export const withAllFunctions = <F extends (...args: never[]) => unknown>(f: F): F =>
  before(f, () => {
    if (all) return
    all = true
    loadGroups(allFunctions)
  })
