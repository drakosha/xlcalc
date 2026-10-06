/** Financial functions; the implementations are in `./financial/`. */
import type { FnDef } from '../args.js'
import { functions as annuity } from './financial/annuity.js'
import { functions as bond } from './financial/bond.js'
import { functions as coupon } from './financial/coupon.js'
import { functions as depreciation } from './financial/depreciation.js'
import { functions as discount } from './financial/discount.js'
import { functions as iterative } from './financial/iterative.js'
import { functions as odd } from './financial/odd.js'

export const functions: Record<string, FnDef> = { ...annuity, ...bond, ...coupon, ...depreciation, ...discount, ...iterative, ...odd }
