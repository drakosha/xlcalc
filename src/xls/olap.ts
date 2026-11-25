import type { Payload } from '../xlsb/records.js'
import { unicode } from './records.js'

/** A hierarchy, measure, named set or KPI stored in an OLAP PivotTable view. */
export interface XlsOlapHierarchy {
  uniqueName: string
  caption: string
  kind: 'hierarchy' | 'measure' | 'set' | 'kpi'
  dimension?: string
  defaultMember?: string
  allMember?: string
  axis?: 'row' | 'column' | 'page' | 'values'
  hidden?: boolean
  attribute?: boolean
  keyAttribute?: boolean
}

/** The saved OLAP hierarchy catalogue and the view's links between fields and hierarchies. */
export interface XlsOlapPivot {
  sheet: string
  name: string
  range: string
  /** The cache's BIFF stream id. */
  cache: number
  /** Entries in file order, with a null slot for an unreadable hierarchy. */
  hierarchies: (XlsOlapHierarchy | null)[]
  /** Entries in pivot field order. The hierarchy and level are zero-based. */
  fields: { hierarchy?: number; level?: number }[]
}

/** SXTH properties shared by cube hierarchies and the measures shown in the field list. */
export function readHierarchy(payload: Payload): XlsOlapHierarchy {
  payload.skip(4)
  const flags = payload.u32()
  const axis = payload.u16()
  payload.skip(12)
  const uniqueName = unicode(payload)
  const caption = unicode(payload)
  const defaultMember = unicode(payload)
  const allMember = unicode(payload)
  const dimension = unicode(payload)
  const kind = flags & 1 ? 'measure' : flags & 32 ? 'set' : flags & 4096 ? 'kpi' : 'hierarchy'
  const hierarchy: XlsOlapHierarchy = { uniqueName, caption, kind }
  if (dimension) hierarchy.dimension = dimension
  if (defaultMember) hierarchy.defaultMember = defaultMember
  if (allMember) hierarchy.allMember = allMember
  if (axis & 1) hierarchy.axis = 'row'
  else if (axis & 2) hierarchy.axis = 'column'
  else if (axis & 4) hierarchy.axis = 'page'
  else if (axis & 8) hierarchy.axis = 'values'
  if (flags & 64) hierarchy.hidden = true
  if (flags & 128) hierarchy.attribute = true
  if (flags & 2048) hierarchy.keyAttribute = true
  return hierarchy
}

/** SXVDTEx relates a pivot field to the hierarchy and cube level it represents. */
export function readOlapField(payload: Payload): XlsOlapPivot['fields'][number] {
  payload.skip(6)
  const hierarchy = payload.i16()
  const level = payload.i32()
  return hierarchy >= 0 ? { hierarchy, ...(level >= 0 ? { level } : {}) } : {}
}
