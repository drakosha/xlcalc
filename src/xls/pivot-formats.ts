/** Differential formats attached to PivotTable areas (SxFormat, SxRule, SxFilt, SxItm, SxDXF). */
import type { PivotArea, PivotAreaReference } from '../contract.js'
import { formatAddress } from '../model.js'
import { pivotAreaXml } from '../xlsx/write/pivot-area.js'
import type { Payload } from '../xlsb/records.js'
import type { Xml } from '../xlsb/xml.js'
import { dxfn } from './conditional.js'
import type { Styles } from './styles.js'

const TYPES = ['none', 'normal', 'data', 'all', 'origin', 'button', 'topRight'] as const
const AXES: Partial<Record<number, PivotArea['axis']>> = { 1: 'rows', 2: 'columns', 4: 'pages', 8: 'values' }
const SUBTOTALS = ['subtotal', 'sum', 'count', 'average', 'max', 'min', 'product', 'countNums', 'stdDev', 'stdDevp', 'var', 'varp'] as const

interface Area {
  area: PivotArea
  count: number
  refs: { reference: PivotAreaReference; count: number }[]
}

function commonArea(flags: number): PivotArea | undefined {
  const type = TYPES[(flags >>> 4) & 15]
  if (!type) return undefined
  return {
    type, dataOnly: Boolean(flags & 0x200), labelOnly: Boolean(flags & 0x400),
    grandRow: Boolean(flags & 0x800), grandColumn: Boolean(flags & 0x1000),
  }
}

function areaOffset(payload: Payload): string | undefined {
  const top = payload.u8(), bottom = payload.u8(), left = payload.u8(), right = payload.u8()
  if (top > bottom || left > right) return undefined
  return formatAddress(top, left) + ':' + formatAddress(bottom, right)
}

function areaOf(payload: Payload): Area | undefined {
  const fieldPosition = payload.u8()
  const field = payload.u8()
  const flags = payload.u16()
  const layout = payload.u16()
  const count = payload.u16()
  const area = commonArea(flags)
  if (!area) return undefined
  area.fieldPosition = fieldPosition
  area.axis = AXES[flags & 15]
  area.cacheIndex = Boolean(flags & 0x4000)
  area.outline = Boolean(layout & 2)
  if (field !== 255) area.field = field === 254 ? 'values' : field
  if (area.dataOnly || area.type === 'all') area.collapsedLevelsAreSubtotals = true
  if (flags & 0x100) {
    area.offset = areaOffset(payload)
    if (!area.offset) return undefined
  }
  return { area, count, refs: [] }
}

function takeItems(payload: Payload, reference: PivotAreaReference): void {
  const items = (reference.items ??= [])
  while (payload.remaining) items.push(payload.u16())
}

function areaMarkup(area: Area | undefined): string | undefined {
  if (!area || area.refs.length !== area.count
    || area.refs.some(({ reference, count }) => (reference.items?.length ?? 0) !== count)) return undefined
  return pivotAreaXml({ ...area.area, references: area.refs.map(({ reference }) => reference) })
}

function referenceOf(payload: Payload): Area['refs'][number] {
  const flags = payload.u32()
  const field = (flags >>> 16) & 1023
  const subtotalFlags = payload.u16()
  const count = payload.u16()
  const subtotals = SUBTOTALS.filter((_, index) => subtotalFlags & (2 << index))
  const reference: PivotAreaReference = { field: field === 1022 ? 'values' : field, selected: Boolean(flags & 0x04000000) }
  if (subtotals.length) reference.subtotals = subtotals
  return { reference, count }
}

interface Format {
  blank: boolean
  size: number
  area?: Area
  dxf?: number
}

export class PivotFormats {
  private readonly formats: Format[] = []
  private active?: Format

  take(type: number, payload: Payload, styles: Styles): void {
    if (type === 0x00fb) {
      const blank = (payload.u16() & 15) === 0
      const size = payload.u16()
      const format: Format = { blank, size }
      this.formats.push(format)
      this.active = format
      return
    }
    const format = this.active
    if (!format) return
    switch (type) {
      case 0x00f0: format.area = areaOf(payload); break
      case 0x00f2: format.area?.refs.push(referenceOf(payload)); break
      case 0x00f5: case 0x003c: {
        const ref = format.area?.refs.at(-1)
        if (!ref || (type === 0x003c && (ref.reference.items?.length ?? 0) >= ref.count)) break
        takeItems(payload, ref.reference)
        break
      }
      case 0x00f4:
        if (payload.remaining === format.size) format.dxf = styles.dxf(dxfn(payload, true))
        this.active = undefined
        break
    }
  }

  /** Writes complete formats and returns whether any records could not be attached to an area. */
  write(xml: Xml): boolean {
    if (!this.formats.length) return false
    let incomplete = false
    xml.add('<formats>')
    for (const { area, blank, size, dxf } of this.formats) {
      const markup = areaMarkup(area)
      if (!markup || (size && dxf === undefined)) {
        incomplete = true
        continue
      }
      const attrs = (blank ? ' action="blank"' : '') + (dxf === undefined ? '' : ` dxfId="${dxf}"`)
      xml.add(`<format${attrs}>${markup}</format>`)
    }
    xml.add('</formats>')
    return incomplete
  }
}

const SCOPES = ['selection', 'data', 'field'] as const
const RANKS = ['', 'all', 'row', 'column'] as const

interface ConditionalFormat {
  scope: string | undefined
  rank: string | undefined
  priority: number
  count: number
  areas: (Area | undefined)[]
}

/** The areas and priority of a PivotTable's conditional format are in the SXAddl classes. */
export class PivotConditionalFormats {
  private readonly formats: ConditionalFormat[] = []
  private active?: ConditionalFormat

  priorities(): number[] { return this.formats.map(({ priority }) => priority) }

  take(owner: number, property: number, payload: Payload): void {
    if (owner === 0x1b && property === 0x35) {
      payload.skip(6)
      const scope = SCOPES[payload.u32()]
      const rank = RANKS[payload.u32()]
      const priority = payload.u32()
      const count = payload.u32()
      const format = { scope, rank, priority, count, areas: [] }
      this.formats.push(format)
      this.active = format
    } else if (owner === 0x1b && property === 0xff) this.active = undefined
    else if (this.active && owner === 0x0c && property === 0x13) {
      payload.skip(6)
      const flags = payload.u32()
      const layout = payload.u16()
      const offset = payload.section(4)
      const count = payload.u32()
      const position = payload.i32()
      const field = payload.i32()
      let area = commonArea(flags)
      if (area) {
        if (position >= 0) area.fieldPosition = position
        if (field >= 0 || field === -2) area.field = field === -2 ? 'values' : field
        area.outline = Boolean(layout & 2)
        area.collapsedLevelsAreSubtotals = Boolean(flags & 0x10000)
        if (flags & 0x100) {
          area.offset = areaOffset(offset)
          if (!area.offset) area = undefined
        }
      }
      this.active.areas.push(area && { area, count, refs: [] })
    } else if (this.active && owner === 0x0d) {
      const area = this.active.areas.at(-1)
      if (!area) return
      if (property === 0x14) {
        payload.skip(6)
        const flags = payload.u16()
        payload.skip(6)
        const field = payload.i32()
        const count = payload.u32()
        if (field >= 0 || field === -2) {
          const reference: PivotAreaReference = { field: field === -2 ? 'values' : field, selected: Boolean(flags & 16) }
          area.refs.push({ reference, count })
        }
      } else if (property === 0x15) {
        const reference = area.refs.at(-1)?.reference
        if (!reference) return
        payload.skip(6)
        takeItems(payload, reference)
      }
    }
  }

  write(xml: Xml): boolean {
    if (!this.formats.length) return false
    let incomplete = false
    xml.add('<conditionalFormats>')
    for (const { scope, rank, priority, count, areas } of this.formats) {
      const markup = areas.map(areaMarkup)
      if (!scope || rank === undefined || !priority || markup.length !== count || markup.includes(undefined)) {
        incomplete = true
        continue
      }
      xml.add(`<conditionalFormat scope="${scope}"${rank ? ` type="${rank}"` : ''} priority="${priority}">`)
      xml.add('<pivotAreas>' + markup.join('') + '</pivotAreas></conditionalFormat>')
    }
    xml.add('</conditionalFormats>')
    return incomplete
  }
}
