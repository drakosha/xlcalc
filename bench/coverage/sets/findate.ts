/**
 * Financial and date functions: cases beyond the coverage measurement (functions formula.js lacks, day-count
 * bases 0–4, frequencies 1/2/4, the 1900 boundary, leap years, month ends, weekend codes, iteration limits).
 * Excel's answers: `npx tsx bench/coverage/run.ts probes findate` → `docs/excel-cases/findate.json`.
 */
import type { Probe } from '../probes.js'
import { XLFN } from '../signatures.js'

/** Sheet F: cash flows, dates and holidays with mixed content. */
export const SHEETS: Record<string, Record<string, number | string | boolean>> = {
  F: {
    // A: flows for IRR/MIRR/NPV; B: flows with text and a boolean
    A1: -1000, A2: 300, A3: 400, A4: 500, A5: 200, A6: 100,
    B1: -1000, B2: '300', B3: true, B4: 500, B5: 'abc', B6: 700,
    // C: holidays with text, a boolean and a blank (C4); D: holidays as numbers
    C1: 39744, C2: 'abc', C3: true, C5: 39750,
    D1: 39744, D2: 39750, D3: 39744, D4: 39751.5,
    // E/F: XIRR flows and dates; G: dates with text; H: values with a boolean
    E1: -100, E2: 110, F1: 39448, F2: 39813,
    E4: -1000, E5: 100, E6: 100, E7: 1100, F4: 39448, F5: 39814, F6: 40179, F7: 40544,
    G1: 39448, G2: '2009-01-01',
    H1: -100, H2: true,
    // I: an annuity of 12 payments with a final balloon
    I1: -5000, I2: 450, I3: 450, I4: 450, I5: 450, I6: 450, I7: 450, I8: 450, I9: 450, I10: 450, I11: 450, I12: 450, I13: 1450,
  },
}

const D = (y: number, m: number, d: number) => `DATE(${y},${m},${d})`

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  const n = new Map<string, number>()
  const add = (fn: string, ...args: (string | number | boolean)[]) => {
    const k = (n.get(fn) ?? 0) + 1
    n.set(fn, k)
    out.push({ id: `${fn}#${k}`, group: fn, fn, formula: `${XLFN[fn] ?? fn}(${args.map((a) => (typeof a === 'boolean' ? (a ? 'TRUE' : 'FALSE') : a)).join(',')})` })
  }

  // ---------- dates: booleans in former ToolPak functions, text and the 1900 boundary ----------
  for (const [f, args] of [
    ['EDATE', ['TRUE', 1]], ['EDATE', [40558, 'TRUE']], ['EOMONTH', ['TRUE', 1]], ['EOMONTH', [40544, 'TRUE']], ['WEEKNUM', ['TRUE']], ['WEEKNUM', [40977, 'TRUE']],
    ['ISOWEEKNUM', ['TRUE']], ['YEARFRAC', ['TRUE', 41120]], ['YEARFRAC', [40909, 41120, 'TRUE']], ['WORKDAY', ['TRUE', 5]], ['WORKDAY', [39722, 'TRUE']],
    ['NETWORKDAYS.INTL', ['TRUE', 38748]], ['WORKDAY.INTL', [40909, 'TRUE', 11]], ['DATEDIF', ['TRUE', 37622, '"Y"']], ['DAYS360', ['TRUE', 40908]], ['WEEKDAY', [39492, 'TRUE']],
  ] as [string, (string | number)[]][])
    add(f, ...args)
  // EDATE and EOMONTH in February 1900 (Excel's calendar has the 29th)
  for (const [s, m] of [[31, 1], [29, 1], [60, 0], [60, 12], [59, 0], [0, 0], [0, -1], [61, -1], [40558, 1.9], [40558, -1.9], [2958465, 1], [2958435, 0]]) add('EDATE', s, m)
  for (const [s, m] of [[31, 0], [60, 0], [0, -1], [2958465, 0], [2958465, 1]]) add('EOMONTH', s, m)
  // DATEDIF: units, case, February ends, leap days
  for (const u of ['"y"', '"md"', '"Ym"', '"yd"', '"MY"', '"DD"']) add('DATEDIF', 36892, 37622, u)
  for (const [a, b] of [[D(2024, 2, 29), D(2025, 3, 1)], [D(2024, 2, 29), D(2025, 2, 28)], [D(2023, 3, 1), D(2024, 2, 29)], [D(2023, 3, 1), D(2024, 3, 1)], [D(2023, 1, 31), D(2023, 3, 1)], [D(2024, 1, 31), D(2024, 3, 1)], [31, 61], [59, 61], [60, 61], [D(2023, 12, 31), D(2024, 1, 30)]])
    for (const u of ['"MD"', '"YD"', '"M"']) add('DATEDIF', a, b, u)
  // DATEVALUE and TIMEVALUE: forms of date and time text
  for (const t of ['"1/31/2025 6 PM"', '"18:30"', '"45688"', '"2025-01-31 25:00"', '"Jan 31, 2025"', '"31-Jan-25"', '"2/29/1900"', '"1/1/1900"', '"12/31/9999"', '" 2025-01-31 "'])
    add('DATEVALUE', t)
  for (const t of ['"2025-01-31"', '"12:00:00.5"', '"6:30:15 PM"', '"0:00"', '"23:59:59"', '"12 AM"', '"1/31/2025 6 PM"', '"24:00"']) add('TIMEVALUE', t)
  // HOUR, MINUTE, SECOND: rounding to a second near .5 and near midnight
  for (const t of ['(10.4996/86400)', '(10.5004/86400)', '(10.5/86400)', '0.99999999', '0.999994', '0.9999942', '(86399.5/86400)', '45688.999999', '2958465.99999']) {
    add('HOUR', t)
    add('MINUTE', t)
    add('SECOND', t)
  }
  // TIME: INT of each part, limits, negative parts
  for (const [h, m, s] of [[12.9, 0, 0], [0, 0, 59.9], [-1, 120, 0], [1, -61, 0], [0, -1, 0], [32767, 0, 0], [32768, 0, 0], [0, 32767, 32767], [0, 0, -0.5], [24, 0, 0], [0, 1440, 0], ['"1.5"', 0, 0]])
    add('TIME', h, m, s)
  // WEEKDAY and WEEKNUM: every return type, invalid ones
  for (const t of [1, 2, 3, 11, 12, 13, 14, 15, 16, 17, 0, 4, 10, 18, 21, 2.9, '"2"']) add('WEEKDAY', 45658, t)
  for (const t of [1, 2, 11, 12, 13, 14, 15, 16, 17, 21, 0, 3, 18, 22]) add('WEEKNUM', 45658, t)
  for (const s of [1, 7, 8, 365, 366, 45657, 45658, 45664, 46022, 2958465]) for (const t of [1, 2, 16, 21]) add('WEEKNUM', s, t)
  for (const s of [1, 2, 7, 8, 365, 366, 44197, 44561, 45293, 45657, 2958465]) add('ISOWEEKNUM', s)
  // DAYS360 and DAYS: February ends, 31sts, European method
  for (const [a, b] of [[D(2024, 2, 29), D(2024, 3, 31)], [D(2023, 2, 28), D(2023, 3, 31)], [D(2024, 2, 28), D(2024, 3, 31)], [D(2024, 1, 31), D(2024, 2, 29)], [D(2024, 3, 31), D(2024, 2, 29)], [D(2023, 2, 28), D(2024, 2, 29)], [60, 91], [59, 91], [D(2024, 3, 30), D(2024, 5, 31)]])
    for (const e of [false, true]) add('DAYS360', a, b, e)
  add('DAYS360', 40573, 40908, '"TRUE"')
  add('DAYS360', 40573, 40908, 2)
  add('DAYS', '"2025-01-31"', '"abc"')
  add('DAYS', 2958465, 0)

  // ---------- YEARFRAC: all bases on February ends, leap years and long spans ----------
  for (const [a, b] of [
    [D(2024, 2, 29), D(2025, 2, 28)], [D(2023, 2, 28), D(2024, 2, 29)], [D(2024, 1, 31), D(2024, 3, 31)], [D(2023, 12, 31), D(2024, 12, 31)], [D(2024, 3, 1), D(2025, 3, 1)],
    [D(2023, 3, 1), D(2024, 3, 1)], [D(2020, 6, 15), D(2027, 1, 10)], [D(2024, 2, 28), D(2024, 2, 29)], [D(2023, 1, 30), D(2023, 2, 28)], [D(2025, 1, 1), D(2025, 1, 1)],
    [1, 100], [1, 400], [30, 70], [60, 400], [D(1999, 12, 31), D(2000, 12, 31)],
  ])
    for (const b2 of [0, 1, 2, 3, 4]) add('YEARFRAC', a, b, b2)
  for (const b2 of [-1, 5, 4.9, '"1"']) add('YEARFRAC', 40909, 41120, b2)

  // ---------- workdays: weekend codes, holidays with mixed content, direction ----------
  for (const w of [1, 2, 3, 4, 5, 6, 7, 11, 12, 13, 14, 15, 16, 17, 0, 8, 10, 18, 1.5, '"0000000"', '"1111110"', '"0000011"', '"1111111"', '"000001"', '"0000021"', '"11"', 'D!A1', 'D!A4', '"TRUE"'])
    add('NETWORKDAYS.INTL', 45658, 45688, w)
  for (const w of [1, 7, 11, 17, '"1010101"', '"0111111"', 'D!A1', 'D!A7', 3.9]) {
    add('WORKDAY.INTL', 45688, 10, w)
    add('WORKDAY.INTL', 45688, -10, w)
  }
  for (const h of ['F!C1:C5', 'F!D1:D4', 'F!C1', '"2008-10-24"', 'TRUE', 'D!A7', 'D!H1:H7']) {
    add('WORKDAY', 39722, 20, h)
    add('NETWORKDAYS', 39722, 39752, h)
  }
  add('WORKDAY', 39752, -20, 'F!D1:D4')
  add('NETWORKDAYS', 39752, 39722, 'F!D1:D4')
  add('WORKDAY', 39724, 0, 'F!D1:D4')
  add('WORKDAY', 39725, 1, 'F!D1:D4')
  add('WORKDAY', 2958463, 5)
  add('WORKDAY', 3, -5)
  add('NETWORKDAYS', 0, 2958465)
  add('WORKDAY.INTL', 45688, 2, '"0000000"', 'F!D1:D4')

  // ---------- annuities ----------
  for (const r of [-2.5, -1, -0.5, -1.5]) {
    add('FV', r, 10, -200, -500, 1)
    add('PV', r, 10, -200, -500, 0)
    add('PMT', r, 10, 10000)
    add('NPER', r, -100, 1000)
  }
  for (const nper of [0, -3, 1.5]) {
    add('PMT', 0.01, nper, 10000)
    add('PV', 0.01, nper, -100)
    add('FV', 0.01, nper, -100)
  }
  add('PMT', 0, 0, 10000)
  add('NPER', 0, 0, 1000)
  add('NPER', 0.01, 0, 1000)
  add('NPER', 0.01, -10, 1000)
  add('NPER', 0, -100, 1000, 500)
  add('NPER', -0.01, -100, 1000)
  for (const [per, nper] of [[1, 1], [2, 1], [1.5, 1], [1.9999, 1], [2.5, 1.5], [2, 1.5], [3, 2.5], [0.5, 3], [37, 36], [36.5, 36]]) {
    add('IPMT', 0.1, per, nper, 1000)
    add('PPMT', 0.1, per, nper, 1000)
  }
  for (const t of [0, 1]) for (const per of [1, 1.5, 2, 12]) {
    add('IPMT', 0.01, per, 12, 1000, 100, t)
    add('PPMT', 0.01, per, 12, 1000, 100, t)
  }
  add('IPMT', 0, 3, 12, 1000)
  add('PPMT', 0, 3, 12, 1000)
  for (const [nper, s, e, t] of [[24.5, 1, 24, 0], [12, 1, 13, 0], [12, 1.5, 2.5, 0], [12, 0.5, 2, 0], [12, 12, 12, 1], [12, 1, 12, 1], [12.9, 1, 12, 0], [12, 1, 12, 0.5], [12, 1, 12, '"1"']]) {
    add('CUMIPMT', 0.01, nper, 1000, s, e, t)
    add('CUMPRINC', 0.01, nper, 1000, s, e, t)
  }
  add('CUMIPMT', 0.01, 12, 1000, 'TRUE', 2, 0)
  add('CUMPRINC', 0.01, 12, 1000, 1, 2, 'TRUE')
  for (const [x, f] of [[1.1, 10], [1.01, 100], [1.5, 1], [-1.25, 8], [3.14159, 7], [1.12, 1000], [0.5, 2.5], [1.5, -1]]) {
    add('DOLLARDE', x, f)
    add('DOLLARFR', x, f)
  }
  add('EFFECT', 0.05, 0.5)
  add('EFFECT', 0.05, 365)
  add('NOMINAL', 0.05, 2.9)
  add('NOMINAL', 0.05, 'D!A2')
  for (const s of ['F!B1:B6', 'D!K1:K8', 'D!D1:D8', '0.1', '"0.1"', 'TRUE', 'F!A1:A6', 'D!E1:E8']) add('FVSCHEDULE', 100, s)
  for (const [n, pv, fv] of [[10, 100, 0], [10, -100, -200], [10, -100, 200], [0.5, 100, 121], [1e-3, 100, 101]]) add('RRI', n, pv, fv)

  // ---------- iterative rates ----------
  for (const g of [-0.99, -0.9, -0.5, -0.3, 0, 0.5, 1, 2, 5, 10]) add('RATE', 48, -200, 8000, 0, 0, g)
  for (const [nper, p, pv, fv, t] of [[360, -500, 100000, 0, 0], [10, -100, 1000, 0, 1], [10, 0, -1000, 2000, 0], [2.5, -100, 200, 0, 0], [12, 100, -1000, 0, 0], [1, -110, 100, 0, 0], [5, 100, 100, 100, 0], [300, -1, 100, 0, 0], [24, -1000, 20000, -5000, 1]])
    add('RATE', nper, p, pv, fv, t)
  for (const g of [-0.9, -0.5, 0, 0.5, 1, 5]) add('IRR', 'F!A1:A6', g)
  for (const r of ['F!B1:B6', 'F!I1:I13', 'D!E1:E8', 'D!I1:I3', 'F!A1:A2', 'F!E1:E2', 'F!H1:H2', '-100', 'D!B1:C8']) add('IRR', r)
  for (const [f, re] of [[-1, 0.1], [0.1, -1], [0.1, 0.1], ['TRUE', 0.1]]) add('MIRR', 'F!A1:A6', f, re)
  add('MIRR', 'F!B1:B6', 0.1, 0.12)
  add('MIRR', 'F!A1', 0.1, 0.12)
  add('MIRR', 'F!E1:E2', 0.1, 0.12)
  for (const g of [-0.9, -0.5, -0.3, -0.2, -0.1, 0, 0.2, 0.5, 1, 2, 5]) add('XIRR', 'D!J1:J5', 'D!AU1:AU5', g)
  for (const g of [0.1, -0.5, 0.5, 2]) add('XIRR', 'F!E1:E2', 'F!F1:F2', g)
  for (const g of [0.1, -0.5, 0.05]) add('XIRR', 'F!E4:E7', 'F!F4:F7', g)
  add('XIRR', 'F!E1:E2', 'F!G1:G2')
  add('XIRR', 'F!H1:H2', 'F!F1:F2')
  add('XIRR', 'F!B1:B2', 'F!F1:F2')
  add('XIRR', 'D!J1:J5', 'D!AU1:AU5', 'TRUE')
  for (const r of [-0.5, 0.1, 1e-9]) add('XNPV', r, 'F!E4:E7', 'F!F4:F7')
  add('XNPV', 0.1, 'F!E1:E2', 'F!G1:G2')
  add('XNPV', 0.1, 'F!H1:H2', 'F!F1:F2')
  add('XNPV', 0.1, 'D!I13:I17', 'D!AU1:AU5')
  add('XNPV', 0.1, 'F!E1:E2', 'D!K1:K2')

  // ---------- depreciation ----------
  for (const args of [[1e6, 1e5, 6, 7], [1e6, 1e5, 6, 7, 12], [1e6, 1e5, 6, 6], [1e6, 1e5, 6, 2.5, 7], [1e6, 1e5, 6.5, 7, 7], [1e6, 1e5, 6.5, 6.9, 7], [1e6, -1, 6, 1], [-1e6, 1e5, 6, 1], [1e6, 1e5, 0, 1], [1e6, 1e5, 6, 0.5], [1e6, 1e5, 6, 1, 12.5], [1e6, 1e5, 2.5, 3, 7], [1e6, 1e6, 6, 2]])
    add('DB', ...args)
  for (const args of [[2400, 300, 10, 11], [2400, 300, 10, 1.5], [2400, 300, 10, 0.5], [2400, 300, 10, 10.5], [2400, 300, 1, 1], [2400, 300, 1.5, 1], [2400, 300, 10, 2, 10], [2400, 300, 10, 3, 1.5], [-2400, 300, 10, 1], [2400, 0, 10, 10]])
    add('DDB', ...args)
  for (const args of [[30000, 7500, 10, 1.5], [30000, 7500, 10.5, 10.5], [30000, 7500, 10, 0.5], [30000, 7500, 10, 'D!A2']]) add('SYD', ...args)
  for (const args of [
    [2400, 300, 3650, 0, 1], [2400, 300, 120, 0, 1], [2400, 300, 10, 0, 1], [2400, 300, 120, 6, 18], [2400, 300, 120, 6, 18, 1.5], [2400, 300, 10, 0, 0.875, 1.5],
    [2400, 300, 10, 0, 10], [2400, 300, 10, 5, 10], [2400, 300, 10, 2.5, 7.5], [2400, 300, 10, 0, 10, 2, true], [2400, 300, 10, 3.5, 9.25, 2, true], [2400, 300, 10, 8, 10, 3],
    [2400, 0, 5, 0, 5, 1], [2400, 300, 10, 0, 11], [2400, 300, 10, 4, 3], [2400, 300, 10, -1, 3], [2400, 300, 10.5, 0, 10.5], [2400, 300, 10, 9.5, 10], [10000, 1000, 5, 1, 4, 1.5, false],
    [2400, 300, 10, 0, 1, 0], [2400, 3000, 10, 0, 5], ['TRUE', 0, 10, 0, 1], [2400, 300, 10, 0, 1, 2, '"abc"'],
  ])
    add('VDB', ...args)
  const purchased = D(2008, 8, 19)
  const firstPer = D(2008, 12, 31)
  for (const [cost, sal, per, rate, basis] of [
    [2400, 300, 1, 0.15, 1], [2400, 300, 0, 0.15, 1], [2400, 300, 2, 0.15, 1], [2400, 300, 3, 0.15, 1], [2400, 300, 4, 0.15, 1], [2400, 300, 5, 0.15, 1], [2400, 300, 6, 0.15, 1],
    [2400, 300, 7, 0.15, 1], [2400, 300, 1, 0.15, 0], [2400, 300, 1, 0.15, 2], [2400, 300, 1, 0.15, 3], [2400, 300, 1, 0.15, 4], [2400, 300, 1, 0.5, 0], [2400, 300, 1, 0.3, 0],
    [2400, 300, 1, 0.25, 0], [2400, 300, 2, 0.2, 0], [2400, 300, 1, 0.1, 0], [2400, 300, 1.5, 0.15, 1], [2400, 3000, 1, 0.15, 1], [2400, 300, 1, 0, 1], [2400, 300, -1, 0.15, 1], [2400, 300, 1, 0.15, 5],
    [10000, 0, 9, 0.12, 0], [10000, 0, 10, 0.12, 0], [10000, 1000, 3, 0.4, 0],
  ] as [number, number, number, number, number][]) {
    add('AMORDEGRC', cost, purchased, firstPer, sal, per, rate, basis)
    add('AMORLINC', cost, purchased, firstPer, sal, per, rate, basis)
  }
  for (const [p, f] of [[D(2008, 12, 31), D(2008, 12, 31)], [D(2009, 1, 15), D(2008, 12, 31)], [D(2008, 1, 1), D(2008, 12, 31)], [D(2007, 12, 31), D(2008, 12, 31)]]) {
    add('AMORDEGRC', 2400, p, f, 300, 1, 0.15, 1)
    add('AMORLINC', 2400, p, f, 300, 1, 0.15, 1)
  }

  // ---------- coupons: all functions on regular, month-end and February schedules ----------
  const coupCases: [string, string, number][] = [
    [D(2011, 1, 25), D(2011, 11, 15), 2], [D(2011, 1, 25), D(2011, 11, 15), 4], [D(2011, 1, 25), D(2011, 11, 15), 1], [D(2011, 3, 15), D(2015, 8, 31), 2], [D(2012, 3, 15), D(2015, 8, 31), 4],
    [D(2011, 3, 1), D(2015, 2, 28), 2], [D(2012, 2, 29), D(2015, 8, 31), 2], [D(2011, 8, 31), D(2015, 2, 28), 2], [D(2011, 5, 15), D(2011, 11, 15), 2], [D(2011, 11, 14), D(2011, 11, 15), 2],
    [D(2008, 2, 15), D(2017, 11, 15), 2], [D(2011, 9, 30), D(2012, 3, 31), 4], [D(2011, 12, 31), D(2012, 6, 30), 2], ['60', D(2011, 11, 15), 2], ['59', D(2011, 11, 15), 2], ['61', D(2011, 11, 15), 2],
  ]
  for (const [s, m, f] of coupCases) for (const basis of [0, 1, 2, 3, 4]) for (const fn of ['COUPDAYBS', 'COUPDAYS', 'COUPDAYSNC']) add(fn, s, m, f, basis)
  for (const [s, m, f] of coupCases) for (const fn of ['COUPNCD', 'COUPNUM', 'COUPPCD']) add(fn, s, m, f, 0)
  for (const fn of ['COUPDAYBS', 'COUPDAYSNC', 'COUPNCD', 'COUPNUM', 'COUPPCD']) {
    add(fn, D(2011, 1, 25), D(2011, 11, 15), 2)
    add(fn, D(2011, 1, 25), D(2011, 11, 15), 3, 1)
    add(fn, D(2011, 1, 25), D(2011, 11, 15), 2.9, 1)
    add(fn, D(2011, 11, 15), D(2011, 1, 25), 2, 1)
    add(fn, D(2011, 1, 25), D(2011, 11, 15), 'TRUE', 1)
    add(fn, '"2011-01-25"', D(2011, 11, 15), 2, 1)
  }

  // ---------- bonds ----------
  const bonds: [string, string, number, number, number, number][] = [
    [D(2008, 2, 15), D(2017, 11, 15), 0.0575, 0.065, 100, 2], [D(2008, 2, 15), D(2016, 11, 15), 0.0575, 0.065, 100, 2], [D(2008, 2, 15), D(2008, 11, 15), 0.0575, 0.065, 100, 2],
    [D(2008, 2, 15), D(2008, 5, 15), 0.0575, 0.065, 100, 4], [D(2011, 3, 15), D(2015, 8, 31), 0.04, 0.05, 105, 2], [D(2012, 2, 29), D(2020, 2, 29), 0.03, 0.08, 100, 1],
    [D(2010, 1, 1), D(2040, 1, 1), 0, 0.05, 100, 2], [D(2010, 1, 1), D(2040, 1, 1), 0.1, 0, 100, 4], [D(2018, 7, 1), D(2048, 1, 1), 0.08, 0.09, 100, 2],
  ]
  for (const [s, m, rate, yld, red, f] of bonds)
    for (const basis of [0, 1, 2, 3, 4]) {
      add('PRICE', s, m, rate, yld, red, f, basis)
      add('DURATION', s, m, rate, yld, f, basis)
      add('MDURATION', s, m, rate, yld, f, basis)
    }
  for (const [s, m, rate, pr, red, f] of [
    [D(2008, 2, 15), D(2016, 11, 15), 0.0575, 95.04287, 100, 2], [D(2008, 2, 15), D(2008, 11, 15), 0.0575, 99.5, 100, 2], [D(2008, 2, 15), D(2008, 5, 15), 0.0575, 99.9, 100, 4],
    [D(2011, 3, 15), D(2015, 8, 31), 0.04, 98, 105, 2], [D(2010, 1, 1), D(2040, 1, 1), 0, 30, 100, 2], [D(2010, 1, 1), D(2040, 1, 1), 0.05, 150, 100, 2], [D(2010, 1, 1), D(2040, 1, 1), 0.05, 400, 100, 2],
    [D(2010, 1, 1), D(2011, 1, 1), 0.05, 120, 100, 1], [D(2018, 7, 1), D(2048, 1, 1), 0.08, 2, 100, 2],
  ] as [string, string, number, number, number, number][])
    for (const basis of [0, 1, 3]) add('YIELD', s, m, rate, pr, red, f, basis)
  for (const fn of ['PRICE', 'YIELD']) {
    add(fn, D(2008, 2, 15), D(2017, 11, 15), -0.01, fn === 'PRICE' ? 0.065 : 95, 100, 2, 0)
    add(fn, D(2008, 2, 15), D(2017, 11, 15), 0.0575, fn === 'PRICE' ? -0.065 : -95, 100, 2, 0)
    add(fn, D(2008, 2, 15), D(2017, 11, 15), 0.0575, fn === 'PRICE' ? 0.065 : 95, 0, 2, 0)
    add(fn, D(2008, 2, 15), D(2017, 11, 15), 0.0575, fn === 'PRICE' ? 0.065 : 95, 100, 3, 0)
    add(fn, D(2008, 2, 15), D(2017, 11, 15), 'TRUE', fn === 'PRICE' ? 0.065 : 95, 100, 2, 0)
  }
  add('DURATION', D(2008, 1, 1), D(2016, 1, 1), -0.08, 0.09, 2, 1)
  add('MDURATION', D(2008, 1, 1), D(2016, 1, 1), 0.08, -0.09, 2, 1)
  add('MDURATION', D(2008, 1, 1), D(2016, 1, 1), 0.08, 0.09, 2, 1)
  // ACCRINT: calc_method, settlement after the first interest date, bases and frequencies
  for (const [iss, first, set] of [[D(2008, 3, 1), D(2008, 8, 31), D(2008, 5, 1)], [D(2008, 3, 1), D(2008, 8, 31), D(2009, 5, 1)], [D(2008, 3, 5), D(2008, 8, 31), D(2010, 1, 15)], [D(2007, 12, 31), D(2008, 6, 30), D(2008, 3, 31)], [D(2008, 1, 15), D(2008, 4, 15), D(2008, 2, 29)]])
    for (const basis of [0, 1, 2, 3, 4]) for (const calc of [true, false]) add('ACCRINT', iss, first, set, 0.1, 1000, 2, basis, calc)
  for (const f of [1, 4]) add('ACCRINT', D(2008, 3, 1), D(2008, 8, 31), D(2009, 5, 1), 0.1, 1000, f, 0)
  add('ACCRINT', D(2008, 3, 1), D(2008, 8, 31), D(2009, 5, 1), 0.1, 1000, 3, 0)
  add('ACCRINT', D(2008, 3, 1), D(2008, 8, 31), D(2009, 5, 1), 0.1, 1000, 2, 5)
  add('ACCRINT', D(2008, 3, 1), D(2008, 8, 31), D(2008, 3, 1), 0.1, 1000, 2, 0)
  add('ACCRINT', D(2008, 3, 1), D(2008, 8, 31), D(2009, 5, 1), 0.1, 1000, 2, 0, 0)

  // ---------- odd first and last periods ----------
  const oddF: [string, string, string, string, number, number, number, number][] = [
    [D(2008, 11, 11), D(2021, 3, 1), D(2008, 10, 15), D(2009, 3, 1), 0.0785, 0.0625, 100, 2], [D(2008, 11, 11), D(2021, 3, 1), D(2008, 10, 15), D(2009, 3, 1), 0.0575, 0.0772, 100, 2],
    [D(2008, 11, 11), D(2021, 3, 1), D(2008, 1, 15), D(2009, 3, 1), 0.0785, 0.0625, 100, 2], [D(2008, 1, 20), D(2021, 3, 1), D(2007, 7, 15), D(2009, 3, 1), 0.0785, 0.0625, 100, 2],
    [D(2008, 11, 11), D(2010, 3, 1), D(2008, 10, 15), D(2009, 3, 1), 0.05, 0.06, 100, 4], [D(2008, 11, 11), D(2009, 9, 1), D(2008, 10, 15), D(2009, 3, 1), 0.05, 0.06, 102, 2],
    [D(2008, 11, 11), D(2019, 3, 1), D(2008, 10, 15), D(2009, 3, 1), 0.05, 0.06, 100, 1], [D(2008, 12, 31), D(2020, 2, 29), D(2008, 11, 30), D(2009, 2, 28), 0.05, 0.06, 100, 2],
  ]
  for (const [s, m, iss, fc, rate, yld, red, f] of oddF) for (const basis of [0, 1, 2, 3, 4]) add('ODDFPRICE', s, m, iss, fc, rate, yld, red, f, basis)
  for (const [s, m, iss, fc, rate, , red, f] of oddF) for (const [pr, basis] of [[84.5, 0], [113.6, 1], [99, 2], [101, 4]] as const) add('ODDFYIELD', s, m, iss, fc, rate, pr, red, f, basis)
  add('ODDFPRICE', D(2008, 11, 11), D(2021, 3, 1), D(2008, 12, 15), D(2009, 3, 1), 0.0785, 0.0625, 100, 2, 1)
  add('ODDFPRICE', D(2008, 11, 11), D(2021, 3, 1), D(2008, 10, 15), D(2008, 11, 1), 0.0785, 0.0625, 100, 2, 1)
  add('ODDFPRICE', D(2008, 11, 11), D(2021, 3, 1), D(2008, 10, 15), D(2009, 3, 1), -0.01, 0.0625, 100, 2, 1)
  const oddL: [string, string, string, number, number, number, number][] = [
    [D(2008, 2, 7), D(2008, 6, 15), D(2007, 10, 15), 0.0375, 0.0405, 100, 2], [D(2008, 4, 20), D(2008, 6, 15), D(2007, 12, 24), 0.0375, 0.0405, 100, 2],
    [D(2008, 2, 7), D(2009, 6, 15), D(2007, 10, 15), 0.0375, 0.0405, 100, 2], [D(2008, 2, 7), D(2008, 6, 15), D(2007, 10, 15), 0.0375, 0.0405, 100, 4],
    [D(2008, 2, 7), D(2008, 12, 15), D(2007, 10, 15), 0.06, 0.05, 100, 1], [D(2008, 2, 29), D(2008, 8, 31), D(2007, 11, 30), 0.05, 0.06, 101, 2],
  ]
  for (const [s, m, last, rate, yld, red, f] of oddL) for (const basis of [0, 1, 2, 3, 4]) add('ODDLPRICE', s, m, last, rate, yld, red, f, basis)
  for (const [s, m, last, rate, , red, f] of oddL) for (const [pr, basis] of [[99.875, 0], [99.5, 1], [100.2, 3], [98, 4]] as const) add('ODDLYIELD', s, m, last, rate, pr, red, f, basis)
  add('ODDLPRICE', D(2008, 2, 7), D(2008, 6, 15), D(2008, 3, 15), 0.0375, 0.0405, 100, 2, 0)
  add('ODDLYIELD', D(2008, 2, 7), D(2008, 6, 15), D(2007, 10, 15), 0.0375, 0, 100, 2, 0)

  // ---------- discounted securities and Treasury bills ----------
  const disc: [string, string][] = [[D(2008, 2, 15), D(2008, 5, 15)], [D(2008, 2, 16), D(2008, 3, 1)], [D(2024, 2, 29), D(2025, 2, 28)], [D(2011, 1, 31), D(2013, 2, 28)], [D(2007, 12, 31), D(2010, 6, 30)]]
  for (const [s, m] of disc)
    for (const basis of [0, 1, 2, 3, 4]) {
      add('INTRATE', s, m, 1000000, 1014420, basis)
      add('RECEIVED', s, m, 1000000, 0.0575, basis)
      add('YIELDDISC', s, m, 99.795, 100, basis)
      add('DISC', s, m, 97.975, 100, basis)
      add('PRICEDISC', s, m, 0.0525, 100, basis)
    }
  add('RECEIVED', D(2008, 2, 15), D(2028, 5, 15), 1000, 0.06, 0)
  add('INTRATE', D(2008, 2, 15), D(2008, 5, 15), 0, 1014420, 2)
  add('YIELDDISC', D(2008, 2, 15), D(2008, 5, 15), 99.795, 0, 2)
  const mats: [string, string, string][] = [[D(2008, 2, 15), D(2008, 4, 13), D(2007, 11, 11)], [D(2008, 3, 15), D(2008, 11, 3), D(2007, 11, 8)], [D(2024, 2, 29), D(2026, 2, 28), D(2023, 2, 28)], [D(2008, 3, 15), D(2010, 11, 3), D(2008, 3, 20)]]
  for (const [s, m, iss] of mats)
    for (const basis of [0, 1, 2, 3, 4]) {
      add('PRICEMAT', s, m, iss, 0.061, 0.061, basis)
      add('YIELDMAT', s, m, iss, 0.0625, 100.0123, basis)
      add('ACCRINTM', iss, s, 0.1, 1000, basis)
    }
  add('ACCRINTM', D(2008, 4, 1), D(2008, 6, 15), 0.1, 1000, 3)
  add('ACCRINTM', D(2008, 4, 1), D(2008, 6, 15), 0.1, 0)
  add('PRICEMAT', D(2008, 2, 15), D(2008, 4, 13), D(2007, 11, 11), -0.061, 0.061, 0)
  add('YIELDMAT', D(2008, 2, 15), D(2008, 4, 13), D(2007, 11, 11), 0.061, 0, 0)
  for (const [s, m] of [[D(2008, 3, 31), D(2008, 12, 1)], [D(2008, 3, 31), D(2009, 3, 31)], [D(2008, 3, 31), D(2009, 4, 1)], [D(2007, 3, 31), D(2008, 3, 31)], [D(2008, 2, 29), D(2009, 2, 28)], [D(2008, 2, 29), D(2009, 3, 1)], [D(2008, 1, 31), D(2008, 7, 31)], [D(2008, 3, 31), D(2008, 9, 29)], [D(2008, 3, 31), D(2008, 9, 30)], [D(2008, 3, 31), D(2008, 10, 1)]]) {
    add('TBILLEQ', s, m, 0.0914)
    add('TBILLPRICE', s, m, 0.09)
    add('TBILLYIELD', s, m, 98.45)
  }
  add('TBILLPRICE', D(2008, 3, 31), D(2008, 6, 1), 10)
  add('TBILLEQ', D(2008, 3, 31), D(2008, 6, 1), 5)
  add('TBILLEQ', D(2008, 3, 31), D(2009, 2, 1), 1.5)
  add('TBILLYIELD', D(2008, 3, 31), D(2008, 6, 1), 200)

  // ---------- second round: rules the first answers left open ----------
  // COUPDAYS on actual/actual around month ends: maturities on days 28–31 and settlements through two years
  for (const m of [D(2015, 8, 31), D(2015, 8, 30), D(2015, 8, 28), D(2015, 2, 28), D(2016, 2, 29), D(2015, 6, 30), D(2015, 5, 31), D(2015, 11, 30), D(2015, 12, 31), D(2015, 9, 30), D(2015, 3, 31)])
    for (const [y, mo, d] of [[2011, 1, 15], [2011, 2, 28], [2011, 3, 15], [2011, 5, 31], [2011, 6, 30], [2011, 8, 31], [2011, 9, 15], [2011, 11, 30], [2011, 12, 31], [2012, 2, 29], [2012, 3, 1], [2012, 4, 30], [2012, 8, 30]])
      for (const f of [1, 2, 4]) add('COUPDAYS', D(y, mo, d), m, f, 1)
  // hours past 23 in time text
  for (const t of ['"25:00"', '"99:59"', '"9999:00"', '"10000:00"', '"24:00:00"', '"23:60"', '"1:60"', '"25:00 PM"', '"13:00 PM"', '"0:00 AM"']) add('TIMEVALUE', t)
  for (const t of ['"2025-01-31 48:00"', '"2025-01-31 24:00"', '"25:00"', '"12/31/9999 23:00"', '"12/31/9999 24:00"']) add('DATEVALUE', t)
  add('HOUR', '"25:30"')
  add('MINUTE', '"25:30"')
  // XIRR: small negative guesses, blank and text dates
  for (const g of [-0.01, -0.001, -1e-9, 1e-9, 0.01]) add('XIRR', 'D!J1:J5', 'D!AU1:AU5', g)
  for (const g of [-0.5, -0.9, -0.69]) add('XIRR', 'F!B1:B2', 'F!F1:F2', g)
  add('XIRR', 'F!E1:E2', 'D!K1:K2')
  add('XIRR', 'D!K1:K2', 'F!F1:F2')
  add('XNPV', 0.1, 'D!K1:K2', 'F!F1:F2')
  add('XNPV', 0.1, 'F!B1:B2', 'F!F1:F2')
  // VDB near the end of life and in fractional periods
  for (const [s, e] of [[9, 10], [9, 9.5], [9.25, 10], [9.75, 10], [8.5, 9], [8.5, 10], [0, 9.5], [5.5, 6], [1.5, 2], [6.5, 7], [7.5, 8], [8, 8.5], [8.9, 9.1]]) add('VDB', 2400, 300, 10, s, e)
  for (const [s, e] of [[9.5, 10], [8.5, 9]]) add('VDB', 2400, 300, 10, s, e, 2, true)
  add('VDB', 2400, 300, 10, 2.5, 3.5, 1.5)
  add('VDB', 2400, 0, 10, 9.5, 10)
  // AMORDEGRC: coefficients by asset life and the last periods
  for (const rate of [0.5, 0.45, 0.4, 0.35, 1 / 3, 0.3, 0.26, 0.25, 0.24, 0.22, 0.21, 0.2, 0.17, 1 / 6, 0.16, 0.15, 0.1]) add('AMORDEGRC', 2400, purchased, firstPer, 300, 1, rate, 0)
  for (let p = 0; p <= 11; p++) add('AMORDEGRC', 10000, purchased, firstPer, 0, p, 0.12, 0)
  for (let p = 0; p <= 6; p++) add('AMORDEGRC', 10000, purchased, firstPer, 1000, p, 0.4, 0)
  for (let p = 0; p <= 8; p++) add('AMORDEGRC', 2400, purchased, firstPer, 300, p, 0.15, 0)
  for (const [p, f] of [[D(2008, 12, 30), D(2008, 12, 31)], [D(2008, 1, 1), D(2008, 12, 31)], [D(2009, 1, 1), D(2009, 12, 31)], [D(2008, 12, 31), D(2009, 12, 31)]])
    for (const basis of [0, 1, 3]) {
      add('AMORDEGRC', 2400, p, f, 300, 1, 0.15, basis)
      add('AMORLINC', 2400, p, f, 300, 0, 0.15, basis)
    }
  // TBILLEQ: the year length past 182 days
  for (const [s, m] of [[D(2008, 1, 15), D(2008, 10, 15)], [D(2007, 6, 1), D(2008, 5, 1)], [D(2007, 3, 15), D(2008, 2, 1)], [D(2009, 3, 31), D(2009, 12, 1)], [D(2009, 1, 1), D(2010, 1, 1)]]) add('TBILLEQ', s, m, 0.0914)
  // MIRR with a rate of −1, NETWORKDAYS.INTL and WORKDAY.INTL weekends as literals
  add('MIRR', 'F!I1:I13', -1, 0.1)
  add('MIRR', 'F!B1:B6', -1, 0.1)
  add('WORKDAY.INTL', 45688, 10, '"1111111"')
  add('NETWORKDAYS.INTL', 45658, 45688, 'TRUE')
  add('WORKDAY.INTL', 45688, 10, 'FALSE')
  return out
}
