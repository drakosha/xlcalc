/** Filter columns follow their cells; sort ranges follow the same reference transform as the surrounding operation. */
import type { FilterCriteria, FilterSort } from '../contract.js'
import { parseRange, rectOf, rectText, shiftRect, type Area, type Rect } from '../structure/area.js'
import type { Transform } from '../structure/rewrite.js'
import { sheetKey } from '../syntax/print.js'

function moved(range: string, sheet: string, t: Transform): Rect | undefined {
  const x = parseRange(range)!
  const b = (i: number) => ({ i, abs: true })
  const area: Area = { kind: 'range', r: { lo: b(x.r1), hi: b(x.r2) }, c: { lo: b(x.c1), hi: b(x.c2) } }
  const out = t.ref({ sheet, area })
  return out === undefined ? x : typeof out === 'object' && sheetKey(out.sheet) === sheetKey(sheet) ? rectOf(out.area) : undefined
}

/** A last sort after references move, removed with its range. */
export function sortAfter(sort: FilterSort | undefined, sheet: string, t: Transform): FilterSort | undefined {
  if (!sort) return undefined
  const range = moved(sort.range, sheet, t)
  if (!range) return undefined
  const conditions = sort.conditions?.flatMap((c) => {
    const x = moved(c.range, sheet, t)
    return x ? [{ ...c, range: rectText(x) }] : []
  })
  return { ...sort, range: rectText(range), ...(conditions ? { conditions } : {}) }
}

/** Criteria after a filter's cells move, with each column still relative to the new left edge. */
export function criteriaAfter(criteria: FilterCriteria | undefined, before: Rect, after: Rect, sheet: string, t: Transform): FilterCriteria | undefined {
  if (!criteria) return undefined
  const columns = criteria.columns?.flatMap((c) => {
    const x = moved(rectText({ ...before, c1: before.c1 + c.column, c2: before.c1 + c.column }), sheet, t)
    const column = x ? x.c1 - after.c1 : -1
    return column >= 0 && column <= after.c2 - after.c1 ? [{ ...c, column }] : []
  })
  const sort = sortAfter(criteria.sort, sheet, t)
  const { sort: _, ...out } = criteria
  return { ...out, ...(columns ? { columns } : {}), ...(sort ? { sort } : {}) }
}

/** A table carried or copied whole keeps its relative columns and translates its stored sort. */
export function offsetCriteria(criteria: FilterCriteria | false | undefined, dr: number, dc: number): FilterCriteria | false | undefined {
  if (!criteria || !criteria.sort) return criteria
  const sort = criteria.sort
  const shift = (range: string) => rectText(shiftRect(parseRange(range)!, dr, dc))
  return { ...criteria, sort: { ...sort, range: shift(sort.range), ...(sort.conditions ? { conditions: sort.conditions.map((c) => ({ ...c, range: shift(c.range) })) } : {}) } }
}
