/**
 * COUPDAYS and VDB held out after \`coupdays-vdb-held\` led to the rules' last parts (COUPDAYS counts from the month's last
 * day only up to settlement; VDB's rate of 1 or more takes all but the salvage): COUPDAYS at frequency 4 between the
 * stepped day and the month's end, around 1900; VDB from fractional starts where the straight line has taken over, by
 * factor, the misses of the first held-out set among them. \`PREDICTED\` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import { heldCases } from './arrays/common.js'
import { e, n } from '../vals.js'

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // COUPDAYS at frequency 4 on settlements between the stepped day and the month's end
  ['COUPDAYS', 'COUPDAYS(DATE(2013,11,28),DATE(2015,8,31),4,1)', n(92)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,11,29),DATE(2015,8,31),4,1)', n(92)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,11,30),DATE(2015,8,31),4,1)', n(90)],
  ['COUPDAYS', 'COUPDAYS(DATE(2014,2,28),DATE(2015,8,31),4,1)', n(89)],
  ['COUPDAYS', 'COUPDAYS(DATE(2014,5,29),DATE(2015,8,31),4,1)', n(92)],
  ['COUPDAYS', 'COUPDAYS(DATE(2014,5,30),DATE(2015,8,31),4,1)', n(92)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,8,29),DATE(2015,8,31),4,1)', n(92)],
  // maturities on the 30th and other month ends
  ['COUPDAYS', 'COUPDAYS(DATE(2013,11,29),DATE(2015,8,30),4,1)', n(92)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,11,29),DATE(2015,11,30),4,1)', n(92)],
  ['COUPDAYS', 'COUPDAYS(DATE(2014,5,30),DATE(2015,8,30),4,1)', n(92)],
  ['COUPDAYS', 'COUPDAYS(DATE(2014,5,30),DATE(2015,11,30),4,1)', n(92)],
  // a leap maturity at frequency 2
  ['COUPDAYS', 'COUPDAYS(DATE(2014,2,28),DATE(2016,2,29),2,1)', n(181)],
  ['COUPDAYS', 'COUPDAYS(DATE(2014,8,29),DATE(2016,2,29),2,1)', n(184)],
  ['COUPDAYS', 'COUPDAYS(DATE(2014,8,30),DATE(2016,2,29),2,1)', n(184)],
  // around 1900
  ['COUPDAYS', 'COUPDAYS(DATE(1899,12,31),DATE(1901,6,30),2,1)', e('NUM')],
  ['COUPDAYS', 'COUPDAYS(DATE(1900,2,1),DATE(1902,8,31),4,1)', n(93)],
  ['COUPDAYS', 'COUPDAYS(DATE(1900,3,1),DATE(1903,2,28),2,1)', n(182)],
  // VDB from fractional starts in the straight-line part, by factor
  ['VDB', 'VDB(2400,300,10,8.5,10,1.5)', n(226.9345640625)],
  ['VDB', 'VDB(2400,300,10,6.25,7.75,1.5)', n(226.9345640625)],
  ['VDB', 'VDB(2400,300,10,7,7.5,1.5)', n(75.6448546875)],
  ['VDB', 'VDB(2400,300,10,7.5,8,1.5)', n(75.6448546875)],
  ['VDB', 'VDB(2400,300,10,8.5,10,1)', n(308.25)],
  ['VDB', 'VDB(2400,300,10,6.25,7.75,1)', n(308.25)],
  ['VDB', 'VDB(2400,300,10,7,7.5,1)', n(102.75)],
  ['VDB', 'VDB(2400,300,10,7.5,8,1)', n(102.75)],
  ['VDB', 'VDB(2400,300,10,8.5,10,2.5)', n(0)],
  ['VDB', 'VDB(2400,300,10,6.25,7.75,2.5)', n(100.282287597656)],
  ['VDB', 'VDB(2400,300,10,7,7.5,2.5)', n(10.1806640625)],
  ['VDB', 'VDB(2400,300,10,7.5,8,2.5)', n(0)],
  ['VDB', 'VDB(2400,300,10,8.5,10,3)', n(0)],
  ['VDB', 'VDB(2400,300,10,6.25,7.75,3)', n(0)],
  ['VDB', 'VDB(2400,300,10,7,7.5,3)', n(0)],
  ['VDB', 'VDB(2400,300,10,7.5,8,3)', n(0)],
  // a rate of 1 or more
  ['VDB', 'VDB(1000,100,3,0.5,1,3)', n(225)],
  ['VDB', 'VDB(1000,100,3,1.5,2,3)', n(0)],
  ['VDB', 'VDB(1000,100,3,0.25,1.25,3)', n(675)],
  ['VDB', 'VDB(1000,100,2,0.5,2,2)', n(450)],
  ['VDB', 'VDB(1000,100,1,0.5,1,2)', n(450)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
