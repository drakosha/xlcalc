/** The group `finance` of `xlcalc/functions` (`docs/function-groups.md`). */
import { functions as bond } from '../functions/financial/bond.js'
import { functions as coupon } from '../functions/financial/coupon.js'
import { functions as depreciation } from '../functions/financial/depreciation.js'
import { functions as discount } from '../functions/financial/discount.js'
import { functions as odd } from '../functions/financial/odd.js'
import { group } from './group.js'

/** The group finance: securities and their coupons, odd periods, discounted securities and depreciation. */
export const finance = /* @__PURE__ */ group('finance', bond, coupon, depreciation, discount, odd)
