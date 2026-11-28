/** Projects worksheet print records from MS-XLSB into the shared SpreadsheetML print reader. */
import { escapeAttr, escapeText } from '../xlsx/xml/text.js'
import type { Payload } from './records.js'
import type { ReadPrintRecord } from './worksheet.js'
import type { Xml } from './xml.js'

const MARGINS = ['left', 'right', 'top', 'bottom', 'header', 'footer']
const OPTIONS = ['horizontalCentered', 'verticalCentered', 'headings', 'gridLines']
const HEADER_FLAGS = ['differentOddEven', 'differentFirst', 'scaleWithDoc', 'alignWithMargins']
const HEADER_TEXT = ['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter']
const PRINT_ERRORS = ['displayed', 'blank', 'dash', 'NA']

function pageSetup(payload: Payload): string {
  const paper = payload.u32()
  const scale = payload.u32()
  const horizontalDpi = payload.u32()
  const verticalDpi = payload.u32()
  const copies = payload.u32()
  const firstPage = payload.i32()
  const width = payload.u32()
  const height = payload.u32()
  const flags = payload.u16()
  const printer = payload.nullableWide()
  payload.done()

  const orientation = flags & 0x40 ? 'default' : flags & 2 ? 'landscape' : 'portrait'
  const comments = flags & 0x20 ? (flags & 0x100 ? 'atEnd' : 'asDisplayed') : 'none'
  const errors = PRINT_ERRORS[(flags >>> 9) & 3]
  let attrs = ` paperSize="${paper}" orientation="${orientation}"`
  if (scale !== 100) attrs += ` scale="${scale}"`
  if (width !== 1) attrs += ` fitToWidth="${width}"`
  if (height !== 1) attrs += ` fitToHeight="${height}"`
  if (flags & 1) attrs += ' pageOrder="overThenDown"'
  if (flags & 8) attrs += ' blackAndWhite="1"'
  if (flags & 16) attrs += ' draft="1"'
  if (comments !== 'none') attrs += ` cellComments="${comments}"`
  if (errors !== 'displayed') attrs += ` errors="${errors}"`
  if (flags & 0x80) attrs += ` useFirstPageNumber="1" firstPageNumber="${firstPage >>> 0}"`
  if (horizontalDpi) attrs += ` horizontalDpi="${horizontalDpi}"`
  if (verticalDpi) attrs += ` verticalDpi="${verticalDpi}"`
  if (copies > 1) attrs += ` copies="${copies}"`
  if (printer) attrs += ` r:id="${escapeAttr(printer)}"`
  return `<pageSetup${attrs}/>`
}

function headerFooter(payload: Payload): string {
  const flags = payload.u16()
  const attrs = HEADER_FLAGS.map((key, index) =>
    ` ${key}="${(flags >>> index) & 1}"`).join('')
  let body = ''
  for (const [index, key] of HEADER_TEXT.entries()) {
    const text = payload.nullableWide()
    const active = index < 2 || (flags & (index < 4 ? 1 : 2))
    if (active && text !== undefined) body += `<${key}>${escapeText(text)}</${key}>`
  }
  payload.done()
  return `<headerFooter${attrs}>${body}</headerFooter>`
}

function pageBreak(payload: Payload): string {
  const index = payload.u32()
  const first = payload.u32()
  const last = payload.u32()
  const manual = payload.u32()
  const pivot = payload.u32()
  payload.done()
  const start = first ? ` min="${first}"` : ''
  const pivotAttr = pivot ? ` pt="${pivot}"` : ''
  return `<brk id="${index}"${start} max="${last}" man="${manual}"${pivotAttr}/>`
}

export function printRecords(xml: Xml): ReadPrintRecord {
  let customView = false
  return (type, payload) => {
    switch (type) {
      case 422: // BrtBeginUserShViews
        customView = true
        return true
      case 425: // BrtEndUserShViews
        customView = false
        return true
      case 476: case 477: case 478: case 479: case 392: case 393: case 394: case 395: case 396: case 552:
        if (customView) return true
        break
      default:
        return false
    }
    switch (type) {
      case 476: {
        const attrs = MARGINS.map((key) => ` ${key}="${payload.f64()}"`).join('')
        payload.done()
        xml.add(`<pageMargins${attrs}/>`)
        break
      }
      case 477: {
        const flags = payload.u16()
        payload.done()
        const attrs = OPTIONS.filter((_, index) => flags & (1 << index)).map((key) => ` ${key}="1"`).join('')
        if (attrs) xml.add(`<printOptions${attrs}/>`)
        break
      }
      case 478: xml.add(pageSetup(payload)); break
      case 479: xml.add(headerFooter(payload)); break
      case 392: case 394: {
        const count = payload.u32()
        const manual = payload.u32()
        payload.done()
        const tag = type === 392 ? 'rowBreaks' : 'colBreaks'
        xml.add(`<${tag} count="${count}" manualBreakCount="${manual}">`)
        break
      }
      case 393: case 395:
        payload.done()
        xml.add(type === 393 ? '</rowBreaks>' : '</colBreaks>')
        break
      case 396: xml.add(pageBreak(payload)); break
      case 552: xml.add('<legacyDrawingHF/>'); break
    }
    return true
  }
}
