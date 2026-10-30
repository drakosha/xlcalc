/** Each group's module by its name, imported where a group is loaded by name (a pool's thread, `worker/thread.ts`), so that a bundle splits them. */
import type { FunctionGroup, FunctionGroupName } from '../contract.js'

export const GROUP_MODULES: Readonly<Record<FunctionGroupName, () => Promise<FunctionGroup>>> = {
  engineering: () => import('./engineering.js').then((m) => m.engineering),
  finance: () => import('./finance.js').then((m) => m.finance),
  lambda: () => import('./lambda.js').then((m) => m.lambda),
  statistics: () => import('./statistics.js').then((m) => m.statistics),
  xml: () => import('./xml.js').then((m) => m.xml),
}
