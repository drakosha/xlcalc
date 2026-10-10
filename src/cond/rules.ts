/**
 * The rules of conditional formats as data: their ranges as rectangles and back, the cell their formulas are written
 * for (the top-left of the first range), and the formulas Excel writes for rules it computes by one (text, blanks and
 * errors, time periods). No computation: the xlsx subpath reads and writes rules with these.
 */
import type { ConditionalFormat } from '../contract.js'
import { formatAddress } from '../model.js'
import { areaText } from '../validation/rules.js'
import { parseRange, type Rect } from '../structure/area.js'

/** The rectangles of a rule's ranges. */
export const rectsOf = (f: ConditionalFormat): Rect[] => f.ranges.map((x) => parseRange(x)!)

/** The ranges of rectangles, `was` (the rule's own spelling) when they are its rectangles. */
export function rangesOf(rects: Rect[], was: string[]): string[] {
  const old = was.map((x) => parseRange(x)!)
  const same = old.length === rects.length && rects.every((x, i) => x.r1 === old[i].r1 && x.r2 === old[i].r2 && x.c1 === old[i].c1 && x.c2 === old[i].c2)
  return same ? was : rects.map(areaText)
}

/**
 * The cell a rule's formulas are written for: the top-left of the rectangle around its ranges, whatever their order
 * (the set `cond/looks`: Excel computes `$A3>2` over `H3:H6 J1:J3` for H1, and saves the ranges as `J1:J3 H3:H6` with
 * the same formula).
 */
export function originOf(f: ConditionalFormat): { r: number; c: number } {
  const xs = rectsOf(f)
  return { r: Math.min(...xs.map((x) => x.r1)), c: Math.min(...xs.map((x) => x.c1)) }
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
