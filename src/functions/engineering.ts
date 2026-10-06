/** Engineering functions, assembled from `./engineering/`. */
import type { FnDef } from '../args.js'
import { functions as bessel } from './engineering/bessel.js'
import { functions as bits } from './engineering/bits.js'
import { functions as complex } from './engineering/complex.js'
import { functions as convert } from './engineering/convert.js'
import { functions as erf } from './engineering/erf.js'
import { functions as radix } from './engineering/radix.js'
import { functions as step } from './engineering/step.js'

export const functions: Record<string, FnDef> = { ...bessel, ...bits, ...complex, ...convert, ...erf, ...radix, ...step }
