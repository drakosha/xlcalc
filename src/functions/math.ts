/** Math and trigonometry functions, assembled from `./math/`. */
import type { FnDef } from '../args.js'
import { functions as arithmetic } from './math/arithmetic.js'
import { functions as combinatorics } from './math/combinatorics.js'
import { functions as matrix } from './math/matrix.js'
import { functions as numerals } from './math/numerals.js'
import { functions as rounding } from './math/rounding.js'
import { functions as sums } from './math/sums.js'
import { functions as trig } from './math/trig.js'

export { power } from './math/arithmetic.js'

export const functions: Record<string, FnDef> = { ...arithmetic, ...combinatorics, ...matrix, ...numerals, ...rounding, ...sums, ...trig }
