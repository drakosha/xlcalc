/** Excel's compatibility names of the descriptive functions: the functions of the other files, under the old names and argument lists. */
import { functions as descriptive } from './descriptive.js'
import type { FnDef } from '../../args.js'
import { functions as regression } from './regression.js'

/** An old name of a newer function takes Excel's 255 arguments, where the newer name takes 254 (`arity` set). */
const older = (def: FnDef): FnDef => ({ ...def, max: undefined })

export const functions: Record<string, FnDef> = {
  COVAR: regression['COVARIANCE.P'],
  MODE: older(descriptive['MODE.SNGL']),
  PERCENTILE: descriptive['PERCENTILE.INC'],
  PERCENTRANK: descriptive['PERCENTRANK.INC'],
  QUARTILE: descriptive['QUARTILE.INC'],
  RANK: descriptive['RANK.EQ'],
  STDEV: older(descriptive['STDEV.S']),
  STDEVP: older(descriptive['STDEV.P']),
  VAR: older(descriptive['VAR.S']),
  VARP: older(descriptive['VAR.P']),
}
