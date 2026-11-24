import { escapeText } from '../xlsx/xml/text.js'
import { range } from './fields.js'
import { NS, type Part } from './parts.js'
import { records } from './records.js'
import { richString } from './strings.js'

export function comments({ bytes, path, xml, context, pending }: Part): void {
  let authors = 0
  for (const { type, payload } of records(bytes, path)) {
    switch (type) {
      case 628: xml.add(`<comments ${NS}>`); break
      case 629: xml.add('</comments>'); break
      case 630: xml.add('<authors>'); break
      case 631: xml.add('</authors>'); break
      case 632:
        xml.add(`<author>${escapeText(payload.wide())}</author>`)
        authors++
        break
      case 633: xml.add('<commentList>'); break
      case 634: xml.add('</commentList>'); break
      case 635: {
        const author = payload.u32()
        if (author >= authors) payload.fail(`comment author index ${author} out of bounds`)
        const ref = range(payload)
        payload.skip(16)
        xml.add(`<comment ref="${ref}" authorId="${author}">`)
        break
      }
      case 636: xml.add('</comment>'); break
      case 637: xml.add(`<text>${richString(payload, context.fonts, pending)}</text>`); break
      default: continue
    }
    payload.done()
  }
}
