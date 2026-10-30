/** Financial functions of the core; the implementations are in `./financial/`. */
import type { FnDef } from '../args.js'
import { functions as annuity } from './financial/annuity.js'
import { functions as iterative } from './financial/iterative.js'

// securities, coupons and depreciation are the group `finance` (`groups/finance.ts`)
export const functions: Record<string, FnDef> = { ...annuity, ...iterative }
