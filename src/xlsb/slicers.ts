import { escapeAttr } from '../xlsx/xml/text.js'
import type { Part } from './parts.js'
import { records } from './records.js'

const CROSS_FILTERS = ['none', 'showItemsWithDataAtTop', 'showItemsWithNoData']
const attr = (key: string, value: string) => ` ${key}="${escapeAttr(value)}"`

export function slicers({ bytes, path, xml }: Part): void {
  for (const { type, payload } of records(bytes, path)) {
    switch (type) {
      case 1115: xml.add('<slicers>'); break
      case 1116: xml.add('</slicers>'); break
      case 1083: {
        const flags = payload.u8()
        const start = payload.u32()
        const columns = payload.u32()
        payload.skip(4) // OLAP hierarchy level
        const height = payload.u32()
        const name = payload.wide()
        const cache = payload.wide()
        const caption = flags & 2 ? attr('caption', payload.wide()) : ''
        const style = flags & 4 ? attr('style', payload.wide()) : ''
        xml.add(`<slicer${attr('name', name)}${attr('cache', cache)}${caption}${style} rowHeight="${height}" startItem="${start}"`
          + ` showCaption="${flags & 1}" lockedPosition="${(flags >>> 3) & 1}"`
          + (columns === 1 ? '' : ` columnCount="${columns}"`) + '/>')
        break
      }
      default: continue
    }
    payload.done()
  }
}

function display(flags: number): string {
  const sortOrder = (flags & 3) === 2 ? 'descending' : 'ascending'
  const cross = CROSS_FILTERS[(flags >>> 2) & 3]
  return ` sortOrder="${sortOrder}" customListSort="${(flags >>> 4) & 1}"` + (cross ? attr('crossFilter', cross) : '')
}

export function slicerCache({ bytes, path, xml }: Part): void {
  for (const { type, payload } of records(bytes, path)) {
    switch (type) {
      case 1077:
        xml.add(`<slicerCacheDefinition${attr('name', payload.wide())}${attr('sourceName', payload.wide())}>`)
        break
      case 1078: xml.add('</slicerCacheDefinition>'); break
      case 1085: {
        const count = payload.u32()
        if (count * 8 > payload.remaining) return payload.fail(`invalid slicer pivot table count ${count}`)
        xml.add('<pivotTables>')
        for (let i = 0; i < count; i++) xml.add(`<pivotTable tabId="${payload.u32()}"${attr('name', payload.wide())}/>`)
        xml.add('</pivotTables>')
        break
      }
      case 1100: {
        payload.skip(4)
        const id = payload.u32()
        const flags = payload.u8()
        xml.add(`<tabular pivotCacheId="${id}" showMissing="${(flags >>> 5) & 1}"${display(flags)}>`)
        break
      }
      case 1101: xml.add('</tabular>'); break
      case 1102: {
        const count = payload.u32()
        if (count * 5 !== payload.remaining) return payload.fail(`invalid slicer cache item count ${count}`)
        xml.add('<items>')
        for (let i = 0; i < count; i++) {
          const index = payload.u32()
          const flags = payload.u8()
          xml.add(`<i x="${index}" s="${flags & 1}" nd="${(flags >>> 1) & 1}"/>`)
        }
        xml.add('</items>')
        break
      }
      case 2077: {
        payload.skip(4)
        const column = payload.u32()
        const table = payload.u32()
        xml.add(`<tableSlicerCache tableId="${table}" column="${column}"${display(payload.u8())}/>`)
        break
      }
      case 2105: // The reader records the presence of the extension, regardless of its OLAP levels.
        xml.add('<slicerCacheHideItemsWithNoData/>')
        continue
      default: continue
    }
    payload.done()
  }
}
