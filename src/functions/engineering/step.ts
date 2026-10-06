/** DELTA and GESTEP: exact comparison of numbers. This engine's own. */
import type { FnDef } from '../../args.js'
import { atp } from '../date/atp.js'

/** Excel reads the second number first: its error or text wins. */
const SECOND_FIRST = { first: [1] }

export const functions: Record<string, FnDef> = {
  DELTA: atp('x x?', (a: number, b = 0) => (a === b ? 1 : 0), SECOND_FIRST),
  GESTEP: atp('x x?', (n: number, step = 0) => (n >= step ? 1 : 0), SECOND_FIRST),
}
