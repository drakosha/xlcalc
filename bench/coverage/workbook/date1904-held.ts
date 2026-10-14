/**
 * Held out: the 1904 date system in combinations the set `date1904` did not have, the answers predicted by its rules
 * and committed before Excel ran (`date1904-held.predicted.json`): dates the former Analysis ToolPak's functions
 * reach before the system's first day, date text around it, the parts of dates in leap 1904, negative values shown
 * by other formats.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { WorkbookCase } from './cases.js'
import type { StoredCell } from '../stored.js'
import { formulaCell as f } from '../stored.js'

const SHOWN: [number, string][] = [
  [-100, 'mmm d, yyyy'],
  [-0.1, 'yyyy'],
  [-1.5, 'd h:mm'],
  [-0.000001, 'h:mm:ss'],
  [59, 'dddd d mmmm yyyy'],
  [-59.75, 'm/d/yy h:mm AM/PM'],
  [1000000, 'yyyy-mm-dd'],
  [-2957003, 'm/d/yyyy'],
  [-2957004, 'm/d/yyyy'],
  [0.999999, 'h:mm'],
]

const FS = [
  'EDATE(30,-1)', 'EOMONTH(30,-1)', 'EDATE(0,-48)', 'EDATE(0,-49)', 'EOMONTH(0,-49)', 'WORKDAY.INTL(5,-10)', 'WORKDAY(10,-5)',
  'COUPPCD(DATE(1904,1,20),DATE(1910,6,15),2)', 'COUPNCD(DATE(1904,1,20),DATE(1910,6,15),1)',
  'TEXT(-1.5,"d h:mm")', '"31-Dec-1903"+0', 'DATEVALUE("1/1/1904 00:00")', 'TIMEVALUE("1/1/1903 6:00")', 'TIMEVALUE("1/1/1904 6:00")',
  'WEEKDAY(-1)', 'YEAR("1/1/1904")', 'DAY("2/29/1904")', 'MONTH(59)', 'ISOWEEKNUM(3)', 'WEEKNUM(2,2)',
  'NETWORKDAYS(DATE(1904,1,1),DATE(1904,1,31))', 'YEARFRAC(0,1461,1)', 'DATEDIF(0,59,"md")', 'EDATE("1/31/1904",1)',
  'COUNTIFS(S!A1:A10,">=1/1/1904",S!A1:A10,"<1/1/1905")', 'SUMPRODUCT((S!A1:A10>=DATE(1904,6,1))*1)',
  'DATE(1904,2,29)', 'DATE(1904,13,1)', 'DATE(1905,-10,1)', 'DATE(1905,-11,1)', 'TEXT(DATE(1904,2,29),"dddd")',
  'S!A1+0', 'TEXT(S!A6,"yyyy-mm-dd hh:mm")', 'HOUR(S!A10)', 'MINUTE(-0.5)', 'SECOND(0.5/86400)',
]

const PREDICTED = JSON.parse(readFileSync(new URL('./date1904-held.predicted.json', import.meta.url), 'utf8')) as Record<string, Record<string, Val>[]>

export function buildWorkbookCases(): WorkbookCase[] {
  const cells: Record<string, StoredCell> = {}
  SHOWN.forEach(([value, format], i) => (cells[`A${i + 1}`] = { value, format }))
  const c: WorkbookCase = {
    id: 'held',
    about: 'the 1904 date system, held out',
    book: { sheets: [{ name: 'S', cells, columns: { A: { width: 40 } } }, { name: 'F', cells: Object.fromEntries(FS.map((x, i) => [`A${i + 1}`, f(x)])) }] },
    file: { date1904: true },
    read: [{ sheet: 'F', range: `A1:A${FS.length}` }],
    show: [{ sheet: 'S', range: `A1:A${SHOWN.length}` }],
  }
  return [PREDICTED[c.id] ? { ...c, predicted: PREDICTED[c.id] } : c]
}
