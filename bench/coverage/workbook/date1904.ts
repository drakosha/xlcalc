/**
 * The 1904 date system (`<workbookPr date1904="1">`) against the 1900 one, on the same book: serial numbers of dates
 * typed in and computed, DATE and the functions that read dates, date text, TODAY and NOW, TEXT and what cells show,
 * negative dates and times, the ends of the range, date differences, the former Analysis ToolPak's date functions.
 */
import type { WorkbookCase } from './cases.js'
import type { StoredCell, StoredSheet } from '../stored.js'

const f = (formula: string): StoredCell => ({ formula })

/** S: values shown by a format in a wide column. */
const SHOWN: [StoredCell['value'], string][] = [
  [0, 'm/d/yyyy'],
  [1, 'm/d/yyyy'],
  [45000, 'm/d/yyyy'],
  [-1, 'm/d/yyyy'],
  [-0.5, 'h:mm'],
  [1.75, 'd h:mm'],
  [-1.25, '[h]:mm'],
  [{ date: '2025-01-31' }, 'yyyy-mm-dd'],
  [{ date: '1904-01-01T06:00:00' }, 'yyyy-mm-dd hh:mm'],
  [2957003, 'm/d/yyyy'],
  [2957004, 'm/d/yyyy'],
  [-1, 'General'],
  [0.5, 'h:mm AM/PM'],
  [-45000, 'm/d/yyyy'],
  [60, 'm/d/yyyy'],
  [1462, 'm/d/yyyy'],
  [-0.75, 'hh:mm:ss'],
  [-2.5, 'dddd'],
  [5, 'dddd, mmmm d, yyyy'],
  [-0.25, '[$-F400]h:mm:ss AM/PM'],
]

function shownSheet(): StoredSheet {
  const cells: Record<string, StoredCell> = {}
  SHOWN.forEach(([value, format], i) => (cells[`A${i + 1}`] = { value, format }))
  return { name: 'S', cells, columns: { A: { width: 40 } } }
}

const D = (y: number, m: number, d: number) => `DATE(${y},${m},${d})`
const BOND = `${D(2025, 1, 15)},${D(2030, 6, 15)}`

const FS = [
  // DATE and the parts of a date
  D(1904, 1, 1), D(1904, 1, 2), D(1903, 12, 31), D(1900, 1, 1), D(1900, 3, 1), D(2025, 1, 31), D(4, 1, 1), D(3, 12, 31), D(9999, 12, 31), D(10000, 1, 1), D(1904, 0, 1), D(1904, 1, 0),
  'YEAR(0)', 'MONTH(0)', 'DAY(0)', 'YEAR(1.5)', 'DAY(60)', 'YEAR(-1)', 'DAY(-0.5)', 'YEAR(2957003)', 'YEAR(2957004)', 'WEEKDAY(0)', 'WEEKDAY(1,2)', 'WEEKNUM(0)', 'ISOWEEKNUM(0)', 'HOUR(-0.25)', 'HOUR(1.25)', `WEEKDAY(${D(1904, 1, 1)})`,
  // date and time text
  'DATEVALUE("1/1/1904")', 'DATEVALUE("1/2/1904")', 'DATEVALUE("12/31/1903")', 'DATEVALUE("2/29/1900")', 'DATEVALUE("3/1/1900")', 'DATEVALUE("1/31/2025")', 'DATEVALUE("2025-01-31")', 'DATEVALUE("31-Jan-2025")',
  'VALUE("1/31/2025")', 'VALUE("1/31/2025 6:00 PM")', '"1/31/2025"+0', '"1/1/1903"+0', '"18:00"+0', 'TIMEVALUE("18:00")', 'VALUE("1/1/1904")', 'YEAR(DATEVALUE("1/31"))-YEAR(TODAY())', 'DAY("1/31/2025")',
  // TEXT
  'TEXT(0,"m/d/yyyy")', 'TEXT(1,"yyyy-mm-dd")', 'TEXT(-1,"m/d/yyyy")', 'TEXT(-0.5,"h:mm")', 'TEXT(-1.25,"[h]:mm")', 'TEXT(1.5,"d h:mm")', 'TEXT(0,"dddd")', 'TEXT(-1,"dddd")', 'TEXT(60,"m/d/yyyy")', 'TEXT(2957003,"m/d/yyyy")', 'TEXT(2957004,"m/d/yyyy")', 'TEXT(-45000,"m/d/yyyy")', 'TEXT(-0.5,"General")', 'TEXT(0.5,"m/d/yyyy h:mm")', 'TEXT(0,"[$-F800]dddd, mmmm dd, yyyy")', 'TEXT(-0.25,"[$-F400]h:mm:ss AM/PM")',
  // differences and calendar arithmetic
  `${D(2025, 1, 31)}-${D(2025, 1, 1)}`, 'DAYS("1/31/2025","1/1/2025")', `DATEDIF(0,${D(2025, 1, 31)},"y")`, 'DATEDIF("1/1/1904","1/31/2025","md")', 'DAYS360(0,365)', 'YEARFRAC(0,366,1)', `YEARFRAC(${D(2024, 1, 1)},${D(2025, 1, 1)},1)`,
  'EDATE(0,1)', `EDATE(${D(2025, 1, 31)},1)`, 'EOMONTH(0,0)', 'EOMONTH(0,-1)', 'EDATE(0,-1)', 'WORKDAY(0,1)', 'WORKDAY(0,-1)', 'NETWORKDAYS(0,10)', 'NETWORKDAYS.INTL(0,10,11)', 'WORKDAY.INTL(0,5,7)', 'WORKDAY(0,10,{3,4})',
  // the former Analysis ToolPak's date functions
  `COUPNCD(${BOND},2)`, `COUPPCD(${BOND},2)`, `COUPDAYS(${BOND},2,1)`, `COUPNUM(${BOND},2)`, `PRICE(${BOND},0.05,0.06,100,2,1)`, `YIELD(${BOND},0.05,98,100,2,1)`,
  `ACCRINT(${D(2024, 1, 1)},${D(2024, 7, 1)},${D(2025, 1, 15)},0.05,1000,2,1)`, `AMORDEGRC(1000,${D(2024, 1, 1)},${D(2024, 12, 31)},100,1,0.15,1)`, `DISC(${D(2025, 1, 15)},${D(2025, 7, 15)},97,100)`, `TBILLPRICE(${D(2025, 1, 15)},${D(2025, 7, 15)},0.05)`,
  `ODDFPRICE(${D(2025, 1, 15)},${D(2030, 6, 15)},${D(2024, 11, 1)},${D(2025, 6, 15)},0.05,0.06,100,2,1)`, 'XNPV(0.1,{-1000,1100},{0,365})', 'XIRR({-1000,1100},{0,365})',
  // criteria with date text
  'COUNTIF(S!A1:A20,">=1/1/1904")', 'COUNTIF(S!A1:A20,"<1/1/2000")', 'SUMIFS(S!A1:A20,S!A1:A20,">=1/1/2025")',
  // the clock and dates typed in
  'TODAY()=DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY()))', 'TODAY()<46000', 'NOW()-TODAY()<1', 'YEAR(TODAY())>=2026', 'S!A8', 'YEAR(S!A8)', 'S!A9', 'TEXT(S!A9,"yyyy-mm-dd hh:mm")',
]

const book = (date1904: boolean): WorkbookCase => ({
  id: date1904 ? 'd1904' : 'd1900',
  about: date1904 ? 'the 1904 date system' : 'the same book in the 1900 date system',
  book: { sheets: [shownSheet(), { name: 'F', cells: Object.fromEntries(FS.map((x, i) => [`A${i + 1}`, f(x)])) }] },
  read: [{ sheet: 'F', range: `A1:A${FS.length}` }],
  show: [{ sheet: 'S', range: `A1:A${SHOWN.length}` }],
  ...(date1904 ? { file: { date1904: true } } : {}),
})

export const buildWorkbookCases = (): WorkbookCase[] => [book(false), book(true)]
