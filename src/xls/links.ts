/** Worksheet hyperlinks ([MS-XLS] HLink and HLinkTooltip, [MS-OSHARED] 2.3.7). */
import { escapeAttr } from '../xlsx/xml/text.js'
import type { Payload } from '../xlsb/records.js'
import type { Xml } from '../xlsb/xml.js'
import { reference } from './records.js'

interface Hyperlink {
  ref: string
  url?: string
  location?: string
}

const terminated = (text: string) => text.split('\0', 1)[0]
const wideString = (payload: Payload) => terminated(payload.wide(payload.u32()))

function moniker(payload: Payload): string | undefined {
  const guid = Array.from(payload.bytes(16), (byte) => byte.toString(16).padStart(2, '0')).join('')
  if (guid === 'e0c9ea79f9bace118c8200aa004ba90b') {
    const size = payload.u32()
    if (size % 2) payload.fail('odd URL moniker size')
    return terminated(payload.wide(size / 2))
  }
  if (guid === '0303000000000000c000000000000046') {
    const parents = payload.u16()
    const ansi = terminated(payload.narrow(payload.u32()))
    payload.skip(2)
    if (payload.u16() !== 0xdead) return undefined
    payload.skip(20)
    const size = payload.u32()
    let path = ansi
    if (size) {
      const extra = payload.section(size)
      const length = extra.u32()
      extra.skip(2)
      if (length % 2) extra.fail('odd Unicode file moniker size')
      path = terminated(extra.wide(length / 2))
    }
    const separator = path.includes('/') && !path.includes('\\') ? '/' : '\\'
    return ('..' + separator).repeat(parents) + path
  }
  return undefined
}

function externalTarget(path: string): string {
  if (/^[A-Za-z]:[\\/]/.test(path)) path = 'file:///' + path.replaceAll('\\', '/')
  else if (path.startsWith('\\\\')) path = 'file:///' + path
  else if (path.startsWith('?')) path = '.' + path
  // The moniker can store literal spaces alongside already escaped URI segments.
  return path.replaceAll(' ', '%20')
}

export class Hyperlinks {
  private readonly links: Hyperlink[] = []
  private readonly tips = new Map<string, string>()

  take(type: number, payload: Payload): boolean | undefined {
    if (type === 0x0800) {
      payload.skip(2)
      const ref = reference(payload)
      this.tips.set(ref, terminated(payload.wide(payload.remaining / 2)))
      return true
    }
    if (type !== 0x01b8) return undefined
    const ref = reference(payload)
    payload.skip(16)
    if (payload.u32() !== 2) return false
    const flags = payload.u32()
    if (flags & 16) wideString(payload)
    if (flags & 128) wideString(payload)
    const monikerPath = flags & 1 ? flags & 256 ? wideString(payload) : moniker(payload) : undefined
    if ((flags & 1) && monikerPath === undefined) return false
    const url = monikerPath === undefined ? undefined : externalTarget(monikerPath)
    const location = flags & 8 ? wideString(payload) : undefined
    this.links.push({ ref, url, location })
    return true
  }

  write(xml: Xml, relate: (id: string, kind: string, path: string, external?: boolean) => void): void {
    if (!this.links.length) return
    xml.add('<hyperlinks>')
    for (const [i, { ref, url, location }] of this.links.entries()) {
      const id = `hyperlink${i}`
      if (url) relate(id, 'hyperlink', url, true)
      const tip = this.tips.get(ref)
      xml.add(`<hyperlink ref="${ref}"${url ? ` r:id="${id}"` : ''}${location ? ` location="${escapeAttr(location)}"` : ''}`
        + (tip ? ` tooltip="${escapeAttr(tip)}"` : '') + '/>')
    }
    xml.add('</hyperlinks>')
  }
}
