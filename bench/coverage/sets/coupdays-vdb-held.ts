/**
 * COUPDAYS and VDB held out: the rules fitted on \`coupdays-vdb\` (COUPDAYS on basis 1 from dates stepped back from
 * maturity each from the one before it; VDB's periods counted from the start), checked where they were not fitted.
 * COUPDAYS: settlements between the stepped day and the month's end, month-end maturities of other months and
 * frequencies, maturities on the 29th and 30th, yearly coupons around 29 February, dates before 1900. VDB: starts in
 * the last periods of the life, other factors, fractional lives, salvage at or above cost, no switch in part periods.
 * \`PREDICTED\`, the engine's answers, was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import { heldCases } from './arrays/common.js'
import { e, n } from '../vals.js'

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // settlements between the stepped day of the month and its end (Excel 16.113: 184 at 30 August 2012 against 181 here)
  ['COUPDAYS', 'COUPDAYS(DATE(2012,8,29),DATE(2015,8,31),2,1)', n(184)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,8,29),DATE(2015,8,31),2,1)', n(184)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,8,30),DATE(2015,8,31),2,1)', n(184)],
  ['COUPDAYS', 'COUPDAYS(DATE(2014,8,30),DATE(2015,8,31),2,1)', n(184)],
  ['COUPDAYS', 'COUPDAYS(DATE(2014,2,28),DATE(2015,8,31),2,1)', n(181)],
  ['COUPDAYS', 'COUPDAYS(DATE(2012,8,29),DATE(2015,8,30),2,1)', n(183)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,8,29),DATE(2015,8,30),2,1)', n(183)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,8,30),DATE(2015,8,30),2,1)', n(181)],
  ['COUPDAYS', 'COUPDAYS(DATE(2014,8,30),DATE(2015,8,30),2,1)', n(181)],
  ['COUPDAYS', 'COUPDAYS(DATE(2014,2,28),DATE(2015,8,30),2,1)', n(181)],
  ['COUPDAYS', 'COUPDAYS(DATE(2012,8,29),DATE(2015,8,28),2,1)', n(181)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,8,29),DATE(2015,8,28),2,1)', n(181)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,8,30),DATE(2015,8,28),2,1)', n(181)],
  ['COUPDAYS', 'COUPDAYS(DATE(2014,8,30),DATE(2015,8,28),2,1)', n(181)],
  ['COUPDAYS', 'COUPDAYS(DATE(2014,2,28),DATE(2015,8,28),2,1)', n(181)],
  ['COUPDAYS', 'COUPDAYS(DATE(2012,8,29),DATE(2015,2,28),2,1)', n(184)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,8,29),DATE(2015,2,28),2,1)', n(184)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,8,30),DATE(2015,2,28),2,1)', n(184)],
  ['COUPDAYS', 'COUPDAYS(DATE(2014,8,30),DATE(2015,2,28),2,1)', n(184)],
  ['COUPDAYS', 'COUPDAYS(DATE(2014,2,28),DATE(2015,2,28),2,1)', n(181)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,8,29),DATE(2015,8,31),4,1)', n(92)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,11,29),DATE(2015,8,31),4,1)', n(91)],
  ['COUPDAYS', 'COUPDAYS(DATE(2014,5,30),DATE(2015,8,31),4,1)', n(92)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,1,31),DATE(2017,5,31),4,1)', n(92)],
  // month-end maturities whose dates meet February at frequency 4, or never meet it at 2
  ['COUPDAYS', 'COUPDAYS(DATE(2016,3,15),DATE(2017,5,31),4,1)', n(90)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,4,30),DATE(2017,5,31),4,1)', n(90)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,7,31),DATE(2017,5,31),4,1)', n(92)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,10,31),DATE(2017,5,31),4,1)', n(92)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,12,15),DATE(2017,5,31),4,1)', n(90)],
  ['COUPDAYS', 'COUPDAYS(DATE(2017,2,28),DATE(2017,5,31),4,1)', n(92)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,1,31),DATE(2016,11,30),4,1)', n(91)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,3,15),DATE(2016,11,30),4,1)', n(91)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,4,30),DATE(2016,11,30),4,1)', n(91)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,7,31),DATE(2016,11,30),4,1)', n(92)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,10,31),DATE(2016,11,30),4,1)', n(91)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,12,15),DATE(2016,11,30),4,1)', e('NUM')],
  ['COUPDAYS', 'COUPDAYS(DATE(2017,2,28),DATE(2016,11,30),4,1)', e('NUM')],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,1,31),DATE(2018,4,30),2,1)', n(182)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,3,15),DATE(2018,4,30),2,1)', n(182)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,4,30),DATE(2018,4,30),2,1)', n(183)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,7,31),DATE(2018,4,30),2,1)', n(183)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,10,31),DATE(2018,4,30),2,1)', n(181)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,12,15),DATE(2018,4,30),2,1)', n(181)],
  ['COUPDAYS', 'COUPDAYS(DATE(2017,2,28),DATE(2018,4,30),2,1)', n(181)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,1,31),DATE(2018,10,31),2,1)', n(182)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,3,15),DATE(2018,10,31),2,1)', n(182)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,4,30),DATE(2018,10,31),2,1)', n(183)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,7,31),DATE(2018,10,31),2,1)', n(183)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,10,31),DATE(2018,10,31),2,1)', n(181)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,12,15),DATE(2018,10,31),2,1)', n(181)],
  ['COUPDAYS', 'COUPDAYS(DATE(2017,2,28),DATE(2018,10,31),2,1)', n(181)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,1,31),DATE(2018,7,31),2,1)', n(182)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,3,15),DATE(2018,7,31),2,1)', n(182)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,4,30),DATE(2018,7,31),2,1)', n(182)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,7,31),DATE(2018,7,31),2,1)', n(184)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,10,31),DATE(2018,7,31),2,1)', n(184)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,12,15),DATE(2018,7,31),2,1)', n(184)],
  ['COUPDAYS', 'COUPDAYS(DATE(2017,2,28),DATE(2018,7,31),2,1)', n(181)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,1,31),DATE(2018,1,31),2,1)', n(182)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,3,15),DATE(2018,1,31),2,1)', n(182)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,4,30),DATE(2018,1,31),2,1)', n(182)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,7,31),DATE(2018,1,31),2,1)', n(184)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,10,31),DATE(2018,1,31),2,1)', n(184)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,12,15),DATE(2018,1,31),2,1)', n(184)],
  ['COUPDAYS', 'COUPDAYS(DATE(2017,2,28),DATE(2018,1,31),2,1)', n(181)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,2,28),DATE(2016,2,29),1,1)', n(365)],
  // yearly coupons around 29 February
  ['COUPDAYS', 'COUPDAYS(DATE(2015,3,1),DATE(2016,2,29),1,1)', n(366)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,2,29),DATE(2020,2,29),1,1)', n(365)],
  ['COUPDAYS', 'COUPDAYS(DATE(2017,6,30),DATE(2020,2,29),1,1)', n(365)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,1,30),DATE(2015,1,30),2,1)', n(181)],
  // maturities on the 29th and 30th that are not months' ends
  ['COUPDAYS', 'COUPDAYS(DATE(2013,1,30),DATE(2015,5,29),4,1)', n(90)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,3,1),DATE(2015,1,30),2,1)', n(181)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,3,1),DATE(2015,5,29),4,1)', n(89)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,9,30),DATE(2015,1,30),2,1)', n(184)],
  ['COUPDAYS', 'COUPDAYS(DATE(2013,9,30),DATE(2015,5,29),4,1)', n(92)],
  ['COUPDAYS', 'COUPDAYS(DATE(1900,1,10),DATE(1901,7,31),2,1)', n(184)],
  // before 1900 and through its February
  ['COUPDAYS', 'COUPDAYS(DATE(1900,3,15),DATE(1902,2,28),4,1)', n(89)],
  ['COUPDAYS', 'COUPDAYS(DATE(1900,2,28),DATE(1901,8,31),2,1)', n(181)],
  ['COUPDAYS', 'COUPDAYS(DATE(1900,1,1),DATE(1900,12,31),1,1)', n(365)],
  ['COUPDAYS', 'COUPDAYS(DATE(1900,1,31),DATE(1900,8,31),2,1)', n(184)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,3,15),DATE(2018,8,31),4,0)', n(90)],
  // the other bases and the argument checks
  ['COUPDAYS', 'COUPDAYS(DATE(2016,3,15),DATE(2018,8,31),4,2)', n(90)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,3,15),DATE(2018,8,31),4,3)', n(91.25)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,3,15),DATE(2018,8,31),4,4)', n(90)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,3,15),DATE(2018,8,31),3,1)', e('NUM')],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,3,15),DATE(2018,8,31),2,5)', e('NUM')],
  ['COUPDAYS', 'COUPDAYS(DATE(2018,8,31),DATE(2018,8,31),2,1)', e('NUM')],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,3,15),DATE(2018,8,31),2,"1")', n(182)],
  ['COUPDAYS', 'COUPDAYS(DATE(2016,3,15),DATE(2018,8,31),2,1.9)', n(182)],
  ['VDB', 'VDB(2400,300,10,9.1,10)', n(14.1120866304001)],
  // VDB: starts in the last periods of the life, where the salvage is reached
  ['VDB', 'VDB(2400,300,10,9.9,10)', n(0)],
  ['VDB', 'VDB(2400,300,10,9.25,9.75)', n(3.00820992000004)],
  ['VDB', 'VDB(2400,300,10,9.5,9.75)', n(0)],
  ['VDB', 'VDB(2400,300,10,9,9.25)', n(5.53063680000002)],
  ['VDB', 'VDB(2400,300,10,9.75,9.9)', n(0)],
  ['VDB', 'VDB(1000,200,5,4.5,5)', n(0)],
  ['VDB', 'VDB(1000,200,5,4.2,4.8)', n(0)],
  ['VDB', 'VDB(1000,200,5,3.5,5)', n(0)],
  ['VDB', 'VDB(2400,500,10,8.5,10)', n(0)],
  ['VDB', 'VDB(2400,500,10,7.25,8.75)', n(0)],
  ['VDB', 'VDB(2400,500,10,9.5,10)', n(0)],
  ['VDB', 'VDB(2400,300,10,8.5,10,1.5)', n(226.9345640625)],
  // factors 1.5 and 3 inside periods, a rate of 1 or more, 0 and below
  ['VDB', 'VDB(2400,300,10,6.25,7.75,1.5)', n(226.9345640625)],
  ['VDB', 'VDB(2400,300,10,2.5,4.5,3)', n(509.796)],
  ['VDB', 'VDB(1000,100,2,0.5,1.5,2)', n(400)],
  ['VDB', 'VDB(1000,100,2,0.25,0.75,4)', n(325)],
  ['VDB', 'VDB(2400,300,10,2.5,3.5,0)', n(210)],
  ['VDB', 'VDB(2400,300,10,0,1,-1)', e('NUM')],
  ['VDB', 'VDB(2400,300,7.5,6.5,7.5)', n(23.4982171742114)],
  // fractional lives and ends, a start equal to its end, salvage at or above cost
  ['VDB', 'VDB(2400,300,7.5,5.25,7.25)', n(175.067311934157)],
  ['VDB', 'VDB(2400,300,10,3,3)', n(0)],
  ['VDB', 'VDB(2400,300,10,3.5,3.5)', n(0)],
  ['VDB', 'VDB(2400,2400,10,2.5,5)', n(0)],
  ['VDB', 'VDB(2400,3000,10,2.5,5)', n(0)],
  ['VDB', 'VDB(2400,3000,10,0,10)', n(-600)],
  ['VDB', 'VDB(2400,300,10,9.25,10,2,TRUE)', n(16.5919104000001)],
  // without the switch inside the last periods
  ['VDB', 'VDB(2400,300,10,8.25,9.75,2,TRUE)', n(76.9898880000001)],
  ['VDB', 'VDB(2400,0,10,9.5,10,2,TRUE)', n(32.21225472)],
  ['VDB', 'VDB(2400,300,0,0,1)', e('NUM')],
  // argument checks
  ['VDB', 'VDB(-2400,300,10,0,1)', e('NUM')],
  ['VDB', 'VDB(2400,-300,10,0,1)', e('NUM')],
  ['VDB', 'VDB(2400,300,10,"a",1)', e('VALUE')],
  ['VDB', 'VDB(2400,300,10,0,1,2,2)', n(480)],
]

/** The predictions Excel answered otherwise: the rules' parts these cases led to are checked by \`coupdays-vdb-held-2\`. */
export const MISSED = ['COUPDAYS#6', 'COUPDAYS#7', 'COUPDAYS#8', 'COUPDAYS#9', 'COUPDAYS#11', 'COUPDAYS#12', 'COUPDAYS#13', 'COUPDAYS#14', 'COUPDAYS#77', 'COUPDAYS#78', 'VDB#13', 'VDB#14', 'VDB#16', 'VDB#17']

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
