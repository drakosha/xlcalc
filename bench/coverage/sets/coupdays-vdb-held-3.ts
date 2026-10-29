/**
 * COUPDAYS and VDB held out after \`coupdays-vdb-held-2\` led to the rules' present form: COUPDAYS steps back from the
 * stepped next date to the first one not after settlement, on months' last days by Excel's own calendar; VDB runs its
 * periods from the fractional part of the start. \`PREDICTED\` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import { heldCases } from './arrays/common.js'
import { n } from '../vals.js'

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // COUPDAYS: the chain back to settlement over several periods, maturities on the 31st, 30th and 29th
  ['COUPDAYS', 'COUPDAYS(DATE(2013,11,29),DATE(2015,5,31),4,1)', n(91)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,11,29),DATE(2016,8,31),4,1)', n(91)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,11,30),DATE(2015,5,31),4,1)', n(90)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,11,30),DATE(2016,8,31),4,1)', n(90)],
  ['COUPDAYS', 'COUPDAYS(DATE(2014,2,28),DATE(2015,5,31),4,1)', n(89)],
  ['COUPDAYS', 'COUPDAYS(DATE(2014,2,28),DATE(2016,8,31),4,1)', n(89)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,8,29),DATE(2015,5,31),4,1)', n(92)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,8,29),DATE(2016,8,31),4,1)', n(92)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,5,30),DATE(2015,5,31),4,1)', n(92)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,5,30),DATE(2016,8,31),4,1)', n(92)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,8,29),DATE(2016,8,29),2,1)', n(184)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,8,29),DATE(2016,2,29),4,1)', n(92)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,8,30),DATE(2016,8,29),2,1)', n(184)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,8,30),DATE(2016,2,29),4,1)', n(92)],
  ['COUPDAYS', 'COUPDAYS(DATE(2014,2,28),DATE(2016,8,29),2,1)', n(181)],
  ['COUPDAYS', 'COUPDAYS(DATE(2014,2,28),DATE(2016,2,29),4,1)', n(89)],
  ['COUPDAYS', 'COUPDAYS(DATE(2012,2,29),DATE(2016,11,30),4,1)', n(90)],
  ['COUPDAYS', 'COUPDAYS(DATE(2012,11,29),DATE(2016,11,30),4,1)', n(91)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,2,28),DATE(2016,11,30),4,1)', n(89)],
  // 1900: a 28 February that is not the month's end in Excel's calendar
  ['COUPDAYS', 'COUPDAYS(DATE(1900,2,1),DATE(1901,2,28),2,1)', n(184)],
  ['COUPDAYS', 'COUPDAYS(DATE(1900,2,1),DATE(1901,2,28),1,1)', n(365)],
  ['COUPDAYS', 'COUPDAYS(DATE(1900,1,15),DATE(1903,5,31),4,1)', n(92)],
  ['COUPDAYS', 'COUPDAYS(DATE(1900,3,2),DATE(1902,11,30),4,1)', n(90)],
  ['COUPDAYS', 'COUPDAYS(DATE(1900,2,28),DATE(1900,8,31),2,1)', n(185)],
  // VDB: fractional starts of other parts, factors 1, 1.25, 1.5, 2.5 and 3, fractional lives, salvage above cost, rates of 1 or more, no switch
  ['VDB', 'VDB(2400,300,10,0.1,1.1,2)', n(470.4)],
  ['VDB', 'VDB(2400,300,10,3.3,5.7,2)', n(474.9656064)],
  ['VDB', 'VDB(2400,300,10,7.7,9.9,2)', n(132.8521728)],
  ['VDB', 'VDB(2400,300,10,2.3,6.6,1)', n(885.50961038961)],
  ['VDB', 'VDB(2400,300,10,4.1,8.9,1.5)', n(767.166467197115)],
  ['VDB', 'VDB(2400,300,10,0.7,9.3,2.5)', n(1680)],
  ['VDB', 'VDB(2400,300,10,5.5,9.5,3)', n(42.8628)],
  ['VDB', 'VDB(2400,300,10,8.3,10,1.25)', n(314.374614514802)],
  ['VDB', 'VDB(2400,300,10,9.1,9.9,1)', n(164.512405063291)],
  ['VDB', 'VDB(2400,300,7.4,3.25,6.75)', n(550.635467045653)],
  ['VDB', 'VDB(2400,300,7.4,6.6,7.4)', n(2.90331342520351)],
  ['VDB', 'VDB(2400,300,7.5,0.5,7.5)', n(1780)],
  ['VDB', 'VDB(2400,300,7.5,3.25,6.75)', n(559.198281920439)],
  ['VDB', 'VDB(5000,500,4.5,1.5,4.5,1.5)', n(2277.77777777778)],
  ['VDB', 'VDB(2400,3000,10,0.5,5)', n(-300)],
  ['VDB', 'VDB(2400,3000,10,2.5,10,1.5)', n(0)],
  ['VDB', 'VDB(1000,100,1.5,0.5,1.5,2)', n(450)],
  ['VDB', 'VDB(1000,100,2,0,2,3)', n(900)],
  ['VDB', 'VDB(1000,100,1,0.25,0.5,2)', n(168.75)],
  ['VDB', 'VDB(2400,300,10,0.3,2.7,2,TRUE)', n(935.04)],
  ['VDB', 'VDB(2400,300,10,8.6,10,1.5,TRUE)', n(122.620736267578)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
