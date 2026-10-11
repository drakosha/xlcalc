/**
 * What functions answer for a reference of several areas where they do not take one, measured against Excel en-US (the
 * set `areas`: a union `(D!B1:B3,D!C1:C3)` and a reference across sheets `Jan:Mar!A1:A3` in every argument that
 * takes a range). Without an entry the rule by kind holds (`args.ts`): lists take each area, an array argument is
 * #VALUE! for a union and #REF! across sheets, any other argument #VALUE!. `at` gives an argument its own errors.
 */
import type { AreaErrors } from '../args.js'
import { functions as database } from './database.js'

type Areas = AreaErrors & { at?: Readonly<Record<number, AreaErrors>> }

/** Lookups: no match in a union (#N/A), #VALUE! across sheets. */
const LOOKUP: Areas = { union: 'NA', cube: 'VALUE' }

/** Functions of a rectangle's size or its lines: #REF! of a union, #VALUE! across sheets. */
const SHAPE: Areas = { union: 'REF', cube: 'VALUE' }

/** Lists that do not take several areas, the same error either way. */
const VALUE: Areas = { union: 'VALUE', cube: 'VALUE' }

export const AREAS: Readonly<Record<string, Areas>> = {
  // the database functions
  ...Object.fromEntries(Object.keys(database).map((f) => [f, VALUE])),
  CHOOSECOLS: SHAPE,
  CHOOSEROWS: SHAPE,
  COLUMN: SHAPE,
  CONCAT: { union: 'VALUE', inside: true },
  COLUMNS: SHAPE,
  COUNTBLANK: { cube: 'VALUE' },
  'F.TEST': { union: 'VALUE', cube: 'REF' },
  FILTER: { ...SHAPE, at: { 2: { union: 'VALUE', cube: 'REF' } } },
  FTEST: { union: 'VALUE', cube: 'REF' },
  GROWTH: SHAPE,
  HLOOKUP: { ...LOOKUP, at: { 2: SHAPE } },
  IMPRODUCT: VALUE,
  IMSUM: VALUE,
  IRR: { union: 'NUM', cube: 'VALUE' },
  LINEST: SHAPE,
  LOGEST: { cube: 'VALUE' },
  LOOKUP,
  MATCH: { ...LOOKUP, at: { 2: SHAPE } },
  NETWORKDAYS: VALUE,
  'NETWORKDAYS.INTL': VALUE,
  PERCENTOF: { cube: 'VALUE' },
  ROW: SHAPE,
  ROWS: SHAPE,
  SERIESSUM: VALUE,
  SORT: { cube: 'REF' },
  SUMPRODUCT: { cube: 'REF' },
  TEXTAFTER: { union: 'VALUE', cube: 'REF' },
  TEXTBEFORE: { union: 'VALUE', cube: 'REF' },
  TRANSPOSE: { cube: 'REF' },
  TREND: SHAPE,
  VLOOKUP: { ...LOOKUP, at: { 2: SHAPE } },
  WORKDAY: VALUE,
  'WORKDAY.INTL': VALUE,
  XIRR: { cube: 'VALUE' },
  XLOOKUP: LOOKUP,
  XMATCH: LOOKUP,
  XNPV: { cube: 'VALUE' },
}
