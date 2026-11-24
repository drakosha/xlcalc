import { escapeAttr } from '../xlsx/xml/text.js'
import type { Part } from './parts.js'
import { records } from './records.js'

export function metadata({ bytes, path, xml, context }: Part): void {
  let block = ''
  let types = 0
  xml.add('<metadata xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">')
  for (const { type, payload } of records(bytes, path)) {
    switch (type) {
      case 334:
        types = payload.u32()
        payload.done()
        xml.add(`<metadataTypes count="${types}">`)
        break
      case 335: {
        payload.u32()
        const version = payload.u32()
        const name = payload.wide()
        payload.done()
        xml.add(`<metadataType name="${escapeAttr(name)}" minSupportedVersion="${version}"/>`)
        break
      }
      case 336: xml.add('</metadataTypes>'); break
      case 337: {
        const count = payload.u32()
        const byCell = payload.u32()
        payload.done()
        block = byCell ? 'cellMetadata' : 'valueMetadata'
        context.metadata[byCell ? 'cell' : 'value'] = count
        xml.add(`<${block} count="${count}">`)
        break
      }
      case 338: xml.add(`</${block}>`); block = ''; break
      case 339: {
        const count = payload.u32()
        const name = payload.wide()
        payload.done()
        xml.add(`<futureMetadata name="${escapeAttr(name)}" count="${count}">`)
        break
      }
      case 340: xml.add('</futureMetadata>'); break
      case 51: {
        const count = payload.u32()
        if (count * 8 !== payload.remaining) payload.fail(`invalid metadata entry count ${count}`)
        xml.add('<bk>')
        for (let i = 0; i < count; i++) {
          const type = payload.u32()
          const value = payload.u32()
          if (type < 1 || type > types) payload.fail(`metadata type index ${type} out of bounds`)
          xml.add(`<rc t="${type}" v="${value}"/>`)
        }
        xml.add('</bk>')
        break
      }
      case 52: xml.add('<bk>'); break
      case 53: xml.add('</bk>'); break
      case 4097: {
        const flags = payload.u16()
        payload.done()
        xml.add(`<xda:dynamicArrayProperties fDynamic="${flags & 1}" fCollapsed="${(flags >>> 1) & 1}"/>`)
        break
      }
      case 5003:
        payload.skip(4)
        xml.add(`<xlrd:rvb i="${payload.u32()}"/>`)
        payload.done()
        break
    }
  }
  xml.add('</metadata>')
}
