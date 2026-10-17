/**
 * Date functions on the cases of IronCalc's tests (`xlsx/tests/calc_tests`, files saved by Excel) that the engine
 * answered otherwise: DAYS360 of moments a hair from midnight, WEEKDAY of a date out of range with a return type
 * that is text or an error, WORKDAY with days a hair under 0 and -1, TIMEVALUE of 86,400 seconds. Then the rules
 * around them: DAYS360 near the half second, WEEKDAY's errors in either argument, WORKDAY's fractional days both
 * ways, minutes and seconds of four and five digits. Answers: `test/excel/cases/date-ironcalc-held.json`;
 * `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import { heldCases } from './arrays/common.js'
import { e, n } from '../vals.js'

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // IronCalc's cases
  ['DAYS360', 'DAYS360(45992.000001,46000.999999)', n(9)],
  ['DAYS360', 'DAYS360(DATE(2025,1,1)+0.0000001,DATE(2025,1,1)+0.9999999)', n(1)],
  ['DAYS360', 'DAYS360(DATE(2025,1,1)+0.9999999,DATE(2025,1,2)+0.0000001)', n(0)],
  ['DAYS360', 'DAYS360(DATE(2025,1,1)+0.9999999,DATE(2025,1,1))', n(-1)],
  ['WEEKDAY', 'WEEKDAY(-1,"abc")', e('VALUE')],
  ['WEEKDAY', 'WEEKDAY(-1,1/0)', e('DIV0')],
  ['WEEKDAY', 'WEEKDAY(-1,NA())', e('NA')],
  ['WORKDAY', 'WORKDAY(DATE(2025,1,10),-0.9999999)', n(45666)],
  ['WORKDAY', 'WORKDAY(DATE(2025,1,10),-1e-9)', n(45666)],
  ['WORKDAY.INTL', 'WORKDAY.INTL(DATE(2025,1,10),-1e-9)', n(45666)],
  ['TIMEVALUE', 'TIMEVALUE("0:0:86400")', e('VALUE')],
  // DAYS360 at the half second before midnight, and of whole dates
  ['DAYS360', 'DAYS360(DATE(2025,1,1),DATE(2025,1,30)+(86399.4/86400))', n(29)],
  ['DAYS360', 'DAYS360(DATE(2025,1,1),DATE(2025,1,30)+(86399.6/86400))', n(30)],
  ['DAYS360', 'DAYS360(DATE(2025,1,1)+0.5,DATE(2025,3,1)+0.5)', n(60)],
  ['DAYS360', 'DAYS360(DATE(2025,1,30)+0.99999999,DATE(2025,2,28),TRUE)', n(28)],
  // WEEKDAY's errors in either argument, a bad type with a good date
  ['WEEKDAY', 'WEEKDAY(NA(),1/0)', e('NA')],
  ['WEEKDAY', 'WEEKDAY("abc",1/0)', e('VALUE')],
  ['WEEKDAY', 'WEEKDAY(-1,5)', e('NUM')],
  ['WEEKDAY', 'WEEKDAY(45667,5)', e('NUM')],
  ['WEEKDAY', 'WEEKDAY(2958466,1)', e('NUM')],
  ['WEEKDAY', 'WEEKDAY(2958466,"x")', e('VALUE')],
  // WORKDAY's days rounded down both ways
  ['WORKDAY', 'WORKDAY(DATE(2025,1,10),0.9999999)', n(45667)],
  ['WORKDAY', 'WORKDAY(DATE(2025,1,10),1.5)', n(45670)],
  ['WORKDAY', 'WORKDAY(DATE(2025,1,10),-1.5)', n(45665)],
  ['WORKDAY', 'WORKDAY(DATE(2025,1,10),-0.5)', n(45666)],
  ['WORKDAY.INTL', 'WORKDAY.INTL(DATE(2025,1,10),-2.0000001,1)', n(45664)],
  // minutes and seconds of four and five digits
  ['TIMEVALUE', 'TIMEVALUE("0:0:9999")', n(0.115729166666667)],
  ['TIMEVALUE', 'TIMEVALUE("0:0:10000")', e('VALUE')],
  ['TIMEVALUE', 'TIMEVALUE("0:9999")', n(0.94375)],
  ['TIMEVALUE', 'TIMEVALUE("0:10000")', e('VALUE')],
  ['TIMEVALUE', 'TIMEVALUE("1:59:9999")', n(0.198368055555556)],
  ['VALUE', 'VALUE("0:0:86400")', e('VALUE')],
  ['VALUE', 'VALUE("9999:0:9999")', e('VALUE')],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
