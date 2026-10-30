/**
 * Every group of functions loaded in each test file (`vitest.config.ts`, `setupFiles`), as the package root has them: the
 * tests that build a state from the engine's own modules calculate every function. The light entry's tests run in a
 * process of their own (`test/function-groups.test.ts`).
 */
import { allFunctions } from '../../src/groups/index.js'
import { loadGroups } from '../../src/groups/registry.js'

loadGroups(allFunctions)
