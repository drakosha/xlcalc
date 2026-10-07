/**
 * Date text in the date arguments of the 1904 system: the former Analysis ToolPak's functions (EDATE, EOMONTH, WORKDAY,
 * NETWORKDAYS, WEEKNUM, YEARFRAC, the coupon and bond functions) against the others (DAY, DATEDIF, DAYS360, ISOWEEKNUM),
 * text typed in a formula and read from a cell, holidays as text.
 */
import type { CellInput } from '../../../src/contract.js'
import type { WorkbookCase } from './cases.js'

const f = (formula: string): CellInput => ({ formula })

const FS = [
  'EDATE("1/31/1904",1)', 'EDATE("1/31/2025",1)', 'EDATE(S!A1,1)', 'EOMONTH("1/31/1904",0)', 'EOMONTH("3/15/2025",0)', 'EOMONTH(S!A2,0)',
  'WORKDAY("1/1/1904",1)', 'WORKDAY("1/30/2025",1)', 'WORKDAY(DATE(2025,1,30),1,"1/31/2025")', 'WORKDAY(DATE(2025,1,30),1,S!A3)', 'WORKDAY.INTL("1/1/2025",1,1)',
  'NETWORKDAYS("1/1/2025","1/31/2025")', 'NETWORKDAYS(DATE(2025,1,1),"1/31/2025")', 'NETWORKDAYS.INTL("1/1/2025",DATE(2025,1,31),1)',
  'WEEKNUM("1/5/1904")', 'WEEKNUM("1/5/2025")', 'YEARFRAC("1/1/2025","7/1/2025",1)', 'YEARFRAC(DATE(2025,1,1),"7/1/2025",1)',
  'COUPNCD("1/15/2025","6/15/2030",2)', 'COUPNUM("1/15/2025","6/15/2030",2)', 'COUPDAYBS("1/15/2025","6/15/2030",2,1)', 'PRICE("1/15/2025","6/15/2030",0.05,0.06,100,2,1)',
  'DISC("1/15/2025","7/15/2025",97,100)', 'ACCRINTM("1/1/2025","7/1/2025",0.05,1000,1)', 'TBILLPRICE("1/15/2025","7/15/2025",0.05)',
  'DAY("1/31/1904")', 'DATEDIF("1/1/2025","3/1/2025","d")', 'DAYS360("1/1/2025","3/1/2025")', 'ISOWEEKNUM("1/5/2025")', 'WEEKDAY("1/5/2025")', 'DAYS("3/1/2025","1/1/2025")',
  'EDATE(" 1/31/2025",1)', 'EDATE("2025-01-31",1)', 'EDATE("31-Jan-2025",1)', 'EDATE("1/31/2025 12:00",1)',
  // computed text, text through references and arrays
  'EDATE("1/31/"&"2025",1)', 'EDATE(T("1/31/2025"),1)', 'EDATE(TEXT(44226,"m/d/yyyy"),1)', 'EDATE(IF(TRUE,"1/31/2025"),1)', 'EDATE(S!A1&"",1)',
  'EDATE(IF(TRUE,S!A1),1)', 'EDATE(INDEX(S!A1:A2,1),1)', 'EDATE({"1/31/2025"},1)', 'EDATE(INDEX({"1/31/2025","x"},1),1)', 'EOMONTH(S!A1:A2,0)', 'SUM(EDATE({"1/31/2025";"2/28/2025"},1))',
]

const book = (date1904: boolean): WorkbookCase => ({
  id: date1904 ? 't1904' : 't1900',
  about: date1904 ? 'date text in date arguments, the 1904 date system' : 'the same in the 1900 date system',
  book: {
    sheets: [
      { name: 'S', cells: { A1: { value: '1/31/2025' }, A2: { value: '2/10/2025' }, A3: { value: '1/31/2025' } } },
      { name: 'F', cells: Object.fromEntries(FS.map((x, i) => [`A${i + 1}`, f(x)])) },
    ],
  },
  read: [{ sheet: 'F', range: `A1:A${FS.length}` }],
  ...(date1904 ? { file: { date1904: true } } : {}),
})

export const buildWorkbookCases = (): WorkbookCase[] => [book(false), book(true)]
