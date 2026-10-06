/** DELTA and GESTEP: exact comparison of numbers. This engine's own. */
import type { FnDef } from '../../args.js'

export const functions: Record<string, FnDef> = {
  DELTA: { args: 'x x?', fn: (a: number, b = 0) => (a === b ? 1 : 0) },
  GESTEP: { args: 'x x?', fn: (n: number, step = 0) => (n >= step ? 1 : 0) },
}
