/**
 * Financial and date functions at edges the earlier sets left open: ACCRINT on basis 1 over several coupon periods,
 * dates at the start of 1900 in the bond functions (serial 0, coupon dates before 1900), AMORDEGRC with a purchase
 * date of 0, XNPV and XIRR with a date of 0, and date text among the holidays of WORKDAY and NETWORKDAYS. Answers:
 * `test/excel/cases/findate-edges.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { Probe, SheetCells } from '../probes.js'
import { heldCases } from './arrays/common.js'
import { n, e } from '../vals.js'

/** Sheet H: holidays as date text, as numbers, as numeric text, and text, a boolean and a blank (A6). */
export const SHEETS: SheetCells = {
  H: { A1: '1/6/2025', A2: '2025-01-07', A3: 45665, A4: 'abc', A5: true, A7: '45663', A8: 'Jan 8, 2025' },
}

/** Function, formula and the answer predicted before Excel ran (undefined: not predicted). */
const CASES: [string, string, Val?][] = [
  // ACCRINT on basis 1 across coupon periods: which period's length divides the days of a part of one
  ['ACCRINT', 'ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2025,1,15),0.05,1000,2,1)', n(51.9230769230769)],
  ['ACCRINT', 'ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2024,8,15),0.05,1000,2,1)'],
  ['ACCRINT', 'ACCRINT(DATE(2023,1,1),DATE(2023,7,1),DATE(2024,1,15),0.05,1000,2,1)'],
  ['ACCRINT', 'ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2025,1,15),0.05,1000,2,1,FALSE)'],
  ['ACCRINT', 'ACCRINT(DATE(2024,1,1),DATE(2024,4,1),DATE(2024,11,15),0.05,1000,4,1)'],
  ['ACCRINT', 'ACCRINT(DATE(2023,3,1),DATE(2024,3,1),DATE(2025,6,1),0.05,1000,1,1)'],
  ['ACCRINT', 'ACCRINT(DATE(2024,2,15),DATE(2024,7,1),DATE(2025,3,1),0.05,1000,2,1)'],
  ['ACCRINT', 'ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2025,9,15),0.05,1000,2,1)'],
  ['ACCRINT', 'ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2025,1,1),0.05,1000,2,1)'],
  ['ACCRINT', 'ACCRINT(DATE(2024,2,29),DATE(2024,8,31),DATE(2025,3,15),0.05,1000,2,1)'],
  ['ACCRINT', 'ACCRINT(DATE(2023,1,1),DATE(2023,7,1),DATE(2024,8,15),0.05,1000,2,1)'],
  ['ACCRINT', 'ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2025,1,15),0.05,1000,2,4)'],
  // the start of 1900: a coupon period that begins before it
  ['ACCRINT', 'ACCRINT(1,39691,39569,0.1,1000,2,0)', e('NUM')],
  ['ACCRINT', 'ACCRINT(100,39691,39569,0.1,1000,2,0)'],
  ['ACCRINT', 'ACCRINT(39508,1,39569,0.1,1000,2,0)'],
  ['ACCRINT', 'ACCRINT(39508,39400,39569,0.1,1000,2,0)'],
  ['ACCRINTM', 'ACCRINTM(1,DATE(2008,6,15),0.1,1000,3)'],
  ['PRICE', 'PRICE(1,DATE(2017,11,15),0.0575,0.065,100,2,0)', e('NUM')],
  ['PRICE', 'PRICE(135,DATE(2017,11,15),0.0575,0.065,100,2,0)', e('NUM')],
  ['PRICE', 'PRICE(136,DATE(2017,11,15),0.0575,0.065,100,2,0)'],
  ['PRICE', 'PRICE(137,DATE(2017,11,15),0.0575,0.065,100,2,0)'],
  ['PRICE', 'PRICE(0,DATE(2017,11,15),0.0575,0.065,100,2,1)', e('NUM')],
  ['COUPNUM', 'COUPNUM(135,DATE(2017,11,15),2,0)', e('NUM')],
  ['DURATION', 'DURATION(1,DATE(2048,1,1),0.08,0.09,2,1)', e('NUM')],
  ['MDURATION', 'MDURATION(0,DATE(2048,1,1),0.08,0.09,2,1)', e('NUM')],
  ['YIELD', 'YIELD(136,DATE(2017,11,15),0.0575,95,100,2,0)'],
  // AMORDEGRC and AMORLINC bought on day 0
  ['AMORDEGRC', 'AMORDEGRC(2400,0,DATE(2008,12,31),300,0,0.15,1)'],
  ['AMORDEGRC', 'AMORDEGRC(2400,0,DATE(2008,12,31),300,2,0.15,1)'],
  ['AMORDEGRC', 'AMORDEGRC(2400,0,DATE(2008,12,31),300,1,0.15,0)'],
  ['AMORDEGRC', 'AMORDEGRC(2400,1,DATE(2008,12,31),300,1,0.15,1)'],
  ['AMORDEGRC', 'AMORDEGRC(2400,0,DATE(2008,12,31),300,3,0.15,1)'],
  ['AMORDEGRC', 'AMORDEGRC(2400,DATE(2008,1,1),DATE(2008,12,31),300,1,0.15,1)'],
  ['AMORLINC', 'AMORLINC(2400,0,DATE(2008,12,31),300,0,0.15,1)'],
  // XNPV and XIRR with a date of 0
  ['XNPV', 'XNPV(0.1,{-1000,600,600},{0,365,730})', e('NUM')],
  ['XNPV', 'XNPV(0.1,{-1000,600,600},{1,366,731})'],
  ['XNPV', 'XNPV(0.1,{-1000,600,600},{40000,0,40100})', e('NUM')],
  ['XNPV', 'XNPV(0.1,{-1000,600},{0.5,365})'],
  ['XIRR', 'XIRR({-1000,600,600},{0,365,730})', e('NUM')],
  ['XIRR', 'XIRR({-1000,600,600},{1,366,731})'],
  // date text among the holidays (2025-01-01 is a Wednesday; H!A1 is 6 January, A2 the 7th, A3 the 8th)
  ['WORKDAY', 'WORKDAY(DATE(2025,1,1),5,H!A1)', n(45666)],
  ['WORKDAY', 'WORKDAY(DATE(2025,1,1),5,H!A1:A2)', n(45667)],
  ['WORKDAY', 'WORKDAY(DATE(2025,1,1),5,H!A1:A3)', n(45670)],
  ['WORKDAY', 'WORKDAY(DATE(2025,1,1),5,H!A4)', e('VALUE')],
  ['WORKDAY', 'WORKDAY(DATE(2025,1,1),5,H!A1:A6)', e('VALUE')],
  ['WORKDAY', 'WORKDAY(DATE(2025,1,1),5,H!A5)'],
  ['WORKDAY', 'WORKDAY(DATE(2025,1,1),5,H!A7)', n(45666)],
  ['WORKDAY', 'WORKDAY(DATE(2025,1,1),5,H!A8)', n(45666)],
  ['WORKDAY', 'WORKDAY(DATE(2025,1,1),5,"1/6/2025")'],
  ['WORKDAY', 'WORKDAY(DATE(2025,1,1),5,{"1/6/2025"})'],
  ['WORKDAY.INTL', 'WORKDAY.INTL(DATE(2025,1,1),5,1,H!A1)', n(45666)],
  ['NETWORKDAYS', 'NETWORKDAYS(DATE(2025,1,1),DATE(2025,1,10),H!A1)', n(7)],
  ['NETWORKDAYS.INTL', 'NETWORKDAYS.INTL(DATE(2025,1,1),DATE(2025,1,10),1,H!A1:A2)', n(6)],
  // second run: settlement on a coupon date by basis, an issue two periods before the first interest, AMORDEGRC's rest
  ['ACCRINT', 'ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2025,1,1),0.05,1000,2,0)', n(50)],
  ['ACCRINT', 'ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2025,1,1),0.05,1000,2,2)', n(25 * (1 + 184 / 180))],
  ['ACCRINT', 'ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2025,1,1),0.05,1000,2,3)', n(25 * (1 + 184 / 182.5))],
  ['ACCRINT', 'ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2025,1,1),0.05,1000,2,4)', n(50)],
  ['ACCRINT', 'ACCRINT(DATE(2024,1,1),DATE(2024,4,1),DATE(2024,10,1),0.05,1000,4,1)', n(12.5 * (2 + 92 / 91))],
  ['ACCRINT', 'ACCRINT(DATE(2023,8,1),DATE(2024,7,1),DATE(2024,8,15),0.05,1000,2,1)', n(25 * (153 / 184 + 1 + 45 / 182))],
  ['ACCRINT', 'ACCRINT(DATE(2023,8,1),DATE(2024,7,1),DATE(2024,3,15),0.05,1000,2,1)', n(25 * (153 / 184 + 74 / 182))],
  ['ACCRINT', 'ACCRINT(58,39691,39569,0.1,1000,2,0)', n(0)],
  ['ACCRINT', 'ACCRINT(60,39691,39569,0.1,1000,2,0)'],
  ['AMORDEGRC', 'AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),2000,1,0.15,1)', n(1035)],
  ['AMORDEGRC', 'AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),2000,2,0.15,1)', n(0)],
  ['AMORDEGRC', 'AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),1500,1,0.15,1)', n(1035)],
  ['AMORDEGRC', 'AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),300,5,0.15,1)', n(158)],
  ['AMORDEGRC', 'AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),2300,0,0.15,1)'],
  ['AMORDEGRC', 'AMORDEGRC(2400,DATE(1950,1,1),DATE(2008,12,31),300,0,0.15,1)', n(2100)],
  ['AMORDEGRC', 'AMORDEGRC(2400,DATE(1950,1,1),DATE(2008,12,31),300,1,0.15,1)', n(113)],
  ['AMORLINC', 'AMORLINC(2400,DATE(1950,1,1),DATE(2008,12,31),300,0,0.15,1)', n(2100)],
  ['AMORLINC', 'AMORLINC(2400,DATE(1950,1,1),DATE(2008,12,31),300,1,0.15,1)', n(0)],
  ['XNPV', 'XNPV(0.1,{-1000,600,600},{0,0,365})', e('NUM')],
  ['WORKDAY', 'WORKDAY(DATE(2025,1,1),5,{"1/6/2025";"abc"})', e('VALUE')],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)

/**
 * The predictions Excel did not confirm: ACCRINT whose issue's period begins before 1900 is 0, not #NUM!; DURATION
 * from day 1 is computed (the day is a coupon date); XIRR takes a date of 0; AMORDEGRC's period that passes the
 * salvage takes its whole amount.
 */
export const MISSED: readonly string[] = ['ACCRINT#13', 'DURATION#1', 'XIRR#1', 'AMORDEGRC#7', 'AMORDEGRC#9']
