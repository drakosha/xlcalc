/**
 * Date and financial functions do not depend on the process time zone: the same formulas computed in child
 * processes in five zones (−8 h to +14 h, with past DST changes) give the same values as in this process.
 */
import { execFileSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { compute, type BookInput, type CellValue } from '../src/index.js'

const FORMULAS = [
  'TODAY()', 'NOW()', 'HOUR(NOW())', 'DATE(2026,10,4)', 'DATEVALUE("2008-03-30")', 'DATEVALUE("10/26/2008")', 'TIMEVALUE("2008-03-30 02:30")',
  'YEAR("2025-01-31")', 'DAY(60)', 'WEEKDAY(0)', 'ISOWEEKNUM("2021-01-01")', 'WEEKNUM(45658,21)', 'EDATE("2008-03-30",1)', 'EOMONTH("10/26/2008",0)',
  'DAYS360("2024-02-29","2025-02-28")', 'DATEDIF("2001-01-31","2001-03-01","MD")', 'YEARFRAC("2008-03-30","2009-10-26",1)', 'NETWORKDAYS.INTL("2008-03-01","2008-11-30",7,"2008-03-30")',
  'WORKDAY("2008-03-28",3,"2008-03-31")', 'XIRR(B1:B2,C1:C2)', 'XNPV(0.1,B1:B2,C1:C2)', 'ACCRINT("2008-03-01","2008-08-31","2009-05-01",0.1,1000,2,1)',
  'PRICE("2008-03-30","2017-10-26",0.0575,0.065,100,2,1)', 'YIELD("2008-03-30","2017-10-26",0.0575,95,100,2,1)', 'DURATION("2008-03-30","2017-10-26",0.08,0.09,2,1)',
  'COUPNCD("2008-03-30","2017-10-26",4,1)', 'ODDFPRICE("2008-11-11","2021-03-01","2008-10-15","2009-03-01",0.0785,0.0625,100,2,1)', 'TBILLEQ("2008-03-30","2008-10-26",0.09)',
  'AMORDEGRC(2400,"2008-03-30","2008-10-26",300,1,0.15,1)', 'PRICEMAT("2008-03-30","2008-10-26","2007-11-11",0.061,0.061,1)',
]

const child = fileURLToPath(new URL('./fixtures/findate-child.ts', import.meta.url))

/** The formulas in column A; flows and their dates in B and C. */
const BOOK: BookInput = {
  sheets: [
    {
      name: 'S',
      cells: {
        ...Object.fromEntries(FORMULAS.map((f, i) => [`A${i + 1}`, { formula: f }])),
        B1: { value: -100 },
        B2: { value: 110 },
        C1: { value: { date: '2008-03-30' } },
        C2: { formula: 'DATE(2009,3,30)' },
      },
    },
  ],
}

function inZone(tz: string): { values: CellValue[]; offset: number } {
  const out = execFileSync(process.execPath, ['--import', 'tsx', child, JSON.stringify(BOOK)], { encoding: 'utf8', env: { ...process.env, TZ: tz } })
  return JSON.parse(out.trim().split('\n').at(-1)!)
}

describe('date and financial functions in different time zones', () => {
  const here = compute(BOOK, { clock: { now: '2026-10-04T23:30:00Z', seed: 7 } }).sheets[0].cells
  const local = Object.keys(BOOK.sheets[0].cells).map((a) => here[a].value)

  it('the clock is read in UTC and no value is an error', () => {
    // values come out rounded to 15 significant digits
    expect(local.slice(0, 4)).toEqual([46299, 46299.9791666667, 23, 46299])
    expect(local.filter((v) => v !== null && typeof v === 'object')).toEqual([])
  })

  it('Los Angeles, Moscow, UTC, Kiritimati and Lord Howe give the same values', () => {
    const zones = ['America/Los_Angeles', 'Europe/Moscow', 'UTC', 'Pacific/Kiritimati', 'Australia/Lord_Howe'].map(inZone)
    expect(zones.map((z) => z.offset)).toEqual([480, -180, 0, -840, -660])
    for (const z of zones) expect(z.values).toEqual(local)
  }, 120_000)
})
