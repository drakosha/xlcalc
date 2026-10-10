/**
 * The rules of conditional formats as data: their ranges as rectangles and back, the cell their formulas are written
 * for (the top-left of the first range), and the formulas Excel writes for rules it computes by one (text, blanks and
 * errors, time periods). No computation: the xlsx subpath reads and writes rules with these.
 */
import type { ConditionalFormat } from '../contract.js'
import { formatAddress, indexToCol, MAX_COLS, MAX_ROWS } from '../model.js'
import { parseRange, type Rect } from '../structure/area.js'

/** The rectangles of a rule's ranges. */
export const rectsOf = (f: ConditionalFormat): Rect[] => f.ranges.map((x) => parseRange(x)!)

/** A rectangle as a rule's range: a cell (`B5`), a rectangle (`B2:D9`), whole columns (`B:D`) or whole rows (`2:9`). */
export function rangeText(x: Rect): string {
  if (x.r1 === 0 && x.r2 === MAX_ROWS - 1) return `${indexToCol(x.c1)}:${indexToCol(x.c2)}`
  if (x.c1 === 0 && x.c2 === MAX_COLS - 1) return `${x.r1 + 1}:${x.r2 + 1}`
  const a = formatAddress(x.r1, x.c1)
  return x.r1 === x.r2 && x.c1 === x.c2 ? a : `${a}:${formatAddress(x.r2, x.c2)}`
}

/** The ranges of rectangles, `was` (the rule's own spelling) when they are its rectangles. */
export function rangesOf(rects: Rect[], was: string[]): string[] {
  const old = was.map((x) => parseRange(x)!)
  const same = old.length === rects.length && rects.every((x, i) => x.r1 === old[i].r1 && x.r2 === old[i].r2 && x.c1 === old[i].c1 && x.c2 === old[i].c2)
  return same ? was : rects.map(rangeText)
}

/** The cell a rule's formulas are written for: its first range's top-left. */
export function originOf(f: ConditionalFormat): { r: number; c: number } {
  const x = parseRange(f.ranges[0])!
  return { r: x.r1, c: x.c1 }
}

const quoted = (text: string) => `"${text.replaceAll('"', '""')}"`

/** The formula Excel writes for a rule it computes by one, for the cell `at` (`A1`): what the rule computes without a formula of its own. */
export function defaultFormula(f: ConditionalFormat, at: string): string | undefined {
  switch (f.type) {
    case 'containsText':
      return `NOT(ISERROR(SEARCH(${quoted(f.text)},${at})))`
    case 'notContainsText':
      return `ISERROR(SEARCH(${quoted(f.text)},${at}))`
    case 'beginsWith':
      return `LEFT(${at},LEN(${quoted(f.text)}))=${quoted(f.text)}`
    case 'endsWith':
      return `RIGHT(${at},LEN(${quoted(f.text)}))=${quoted(f.text)}`
    case 'containsBlanks':
      return `LEN(TRIM(${at}))=0`
    case 'notContainsBlanks':
      return `LEN(TRIM(${at}))>0`
    case 'containsErrors':
      return `ISERROR(${at})`
    case 'notContainsErrors':
      return `NOT(ISERROR(${at}))`
    case 'timePeriod': {
      const day = `FLOOR(${at},1)`
      const down = `ROUNDDOWN(${at},0)`
      const month = (d: string) => `AND(MONTH(${at})=MONTH(${d}),YEAR(${at})=YEAR(${d}))`
      switch (f.period) {
        case 'today':
          return `${day}=TODAY()`
        case 'yesterday':
          return `${day}=TODAY()-1`
        case 'tomorrow':
          return `${day}=TODAY()+1`
        case 'last7Days':
          return `AND(TODAY()-${day}<=6,${day}<=TODAY())`
        case 'thisMonth':
          return month('TODAY()')
        case 'lastMonth':
          return month('EDATE(TODAY(),0-1)')
        case 'nextMonth':
          return month('EDATE(TODAY(),0+1)')
        case 'thisWeek':
          return `AND(TODAY()-${down}<=WEEKDAY(TODAY())-1,${down}-TODAY()<=7-WEEKDAY(TODAY()))`
        case 'lastWeek':
          return `AND(TODAY()-${down}>=(WEEKDAY(TODAY())),TODAY()-${down}<(WEEKDAY(TODAY())+7))`
        case 'nextWeek':
          return `AND(${down}-TODAY()>(7-WEEKDAY(TODAY())),${down}-TODAY()<(15-WEEKDAY(TODAY())))`
      }
    }
  }
  return undefined
}

/** The formula a rule computes by, given or Excel's (undefined for a rule computed otherwise). */
export function ruleFormula(f: ConditionalFormat): string | undefined {
  if (f.type === 'expression') return f.formula
  if ('formula' in f && f.formula !== undefined) return f.formula
  const o = originOf(f)
  return defaultFormula(f, formatAddress(o.r, o.c))
}
