/** Projects BIFF print records into the worksheet parts the opt-in SpreadsheetML reader knows. */
import { MAX_COLS, MAX_ROWS } from '../model.js'
import { DEFAULT_MARGINS, type PrintReadIssue } from '../xlsx/print/types.js'
import { escapeText } from '../xlsx/xml/text.js'
import type { Payload } from '../xlsb/records.js'
import type { LegacyEncoding } from './encoding.js'
import { setupFields, type PrintMarkup } from './printing.js'
import { unicode, type RecordInfo } from './records.js'

type Attrs = Record<string, string | number>

function leaf(name: string, attrs: Attrs): string {
  const values = Object.entries(attrs).map(([key, value]) => ` ${key}="${value}"`).join('')
  return values ? `<${name}${values}/>` : ''
}

function text(tag: string, value: string): string {
  return value ? `<${tag}>${escapeText(value)}</${tag}>` : ''
}

/** HeaderFooter carries only the even/first variants. Header and Footer carry the regular text. */
function headerVariants(payload: Payload): { flags: number; body: string } | undefined {
  payload.skip(12)
  const guid = payload.bytes(16)
  if (guid.some(Boolean)) return undefined
  const flags = payload.u16()
  const lengths = [payload.u16(), payload.u16(), payload.u16(), payload.u16()]
  const tags = ['evenHeader', 'evenFooter', 'firstHeader', 'firstFooter']
  let body = ''
  for (const [index, length] of lengths.entries()) {
    if (length === 0) continue
    const value = unicode(payload)
    if (value.length !== length || length > 255) payload.fail('invalid header or footer length')
    body += text(tags[index], value)
  }
  return { flags, body }
}

function pageBreaks(payload: Payload, rows: boolean, legacy: boolean): string {
  const count = payload.u16()
  const tag = rows ? 'rowBreaks' : 'colBreaks'
  const across = rows ? MAX_COLS : MAX_ROWS
  const binaryEnd = rows ? 255 : 65535
  const unbounded = legacy && payload.remaining === count * 2
  let result = ''
  for (let index = 0; index < count; index++) {
    const id = payload.u16()
    const start = unbounded ? 0 : payload.u16()
    const end = unbounded ? binaryEnd : payload.u16()
    // Excel extends whole-width/height BIFF breaks to the edge of the larger XLSX grid.
    const max = start === 0 && end === binaryEnd ? across - 1 : end
    result += leaf('brk', { id, ...(start ? { min: start } : {}), max, man: 1 })
  }
  return result ? `<${tag}>${result}</${tag}>` : ''
}

function pageSetup(payload: Payload): { attrs: Attrs; fit: Attrs; margins: Attrs } {
  const { paper, scale, first, width, height, flags } = setupFields(payload)
  const attrs: Attrs = {}
  if (flags & 1) attrs.pageOrder = 'overThenDown'
  if (flags & 8) attrs.blackAndWhite = 1
  if (flags & 16) attrs.draft = 1
  if (flags & 32) attrs.cellComments = flags & 512 ? 'atEnd' : 'asDisplayed'
  if (flags & 128) {
    attrs.useFirstPageNumber = 1
    attrs.firstPageNumber = first >>> 0
  }
  const errorMode = (flags >>> 10) & 3
  if (errorMode) attrs.errors = ['displayed', 'blank', 'dash', 'NA'][errorMode]

  // When fNoPls is set, these fields are undefined even if Excel wrote nonzero bytes.
  const configured = !(flags & 4)
  if (configured) {
    attrs.paperSize = paper
    if (scale !== 100) attrs.scale = scale
    if (!(flags & 64)) attrs.orientation = flags & 2 ? 'portrait' : 'landscape'
  }

  const margins: Attrs = {}
  if (payload.remaining >= 22) {
    const horizontalDpi = payload.u16()
    const verticalDpi = payload.u16()
    margins.header = payload.f64()
    margins.footer = payload.f64()
    const copies = payload.u16()
    if (configured) {
      if (horizontalDpi) attrs.horizontalDpi = horizontalDpi
      if (verticalDpi) attrs.verticalDpi = verticalDpi
      if (copies > 1) attrs.copies = copies
    }
  }
  return { attrs, fit: { fitToWidth: width, fitToHeight: height }, margins }
}

const MARGIN_NAMES = ['left', 'right', 'top', 'bottom']
const OPTION_NAMES = new Map([[0x002a, 'headings'], [0x002b, 'gridLines'], [0x0083, 'horizontalCentered'], [0x0084, 'verticalCentered']])

export function printMarkup(
  records: Iterable<RecordInfo>, encoding: LegacyEncoding | undefined, report: (part: string, reason: PrintReadIssue['reason']) => void,
): PrintMarkup {
  const options: Attrs = {}
  const margins: Attrs = {}
  let fitToPage = false
  let autoPageBreaks = true
  let setup: ReturnType<typeof pageSetup> | undefined
  let oddHeader = ''
  let oddFooter = ''
  let variants: ReturnType<typeof headerVariants>
  let rowBreaks = ''
  let columnBreaks = ''
  let customView = false

  for (const { type, payload } of records) {
    if (type === 0x01aa) customView = true
    if (type === 0x01ab) customView = false
    if (customView) continue
    let part = 'printOptions'
    try {
      const option = OPTION_NAMES.get(type)
      if (option) {
        if (payload.u16() & 1) options[option] = 1
        continue
      }
      switch (type) {
        case 0x0081: {
          const flags = payload.u16()
          fitToPage = Boolean(flags & 256)
          autoPageBreaks = Boolean(flags & 1)
          break
        }
        case 0x0014:
        case 0x0015: {
          part = 'headerFooter'
          if (!payload.remaining) break
          const value = encoding ? encoding.text(payload, payload.u8()) : unicode(payload)
          if (type === 0x0014) oddHeader = text('oddHeader', value)
          else oddFooter = text('oddFooter', value)
          break
        }
        case 0x089c:
          part = 'headerFooter'
          variants = headerVariants(payload)
          break
        case 0x0866: report('legacyDrawingHF', 'headerFooterDrawing'); break
        case 0x001b:
          part = 'rowBreaks'
          rowBreaks = pageBreaks(payload, true, encoding !== undefined)
          break
        case 0x001a:
          part = 'colBreaks'
          columnBreaks = pageBreaks(payload, false, encoding !== undefined)
          break
        case 0x0026: case 0x0027: case 0x0028: case 0x0029:
          part = 'pageMargins'
          margins[MARGIN_NAMES[type - 0x0026]] = payload.f64()
          break
        case 0x004d: report('pageSetup', 'printerSettings'); break
        case 0x00a1:
          part = 'pageSetup'
          setup = pageSetup(payload)
          break
      }
    } catch {
      report(part, 'invalid')
    }
  }

  const properties = leaf('pageSetUpPr', {
    ...(fitToPage ? { fitToPage: 1 } : {}),
    ...(!autoPageBreaks ? { autoPageBreaks: 0 } : {}),
  })
  const flags = variants?.flags ?? 12
  const headerAttrs = (flags & 1 ? ' differentOddEven="1"' : '') + (flags & 2 ? ' differentFirst="1"' : '')
    + (flags & 4 ? '' : ' scaleWithDoc="0"') + (flags & 8 ? '' : ' alignWithMargins="0"')
  const headerBody = oddHeader + oddFooter + (variants?.body ?? '')
  const header = headerAttrs || headerBody ? `<headerFooter${headerAttrs}>${headerBody}</headerFooter>` : ''
  if (options.gridLines) options.gridLinesSet = 1
  const marginValues = { ...margins, ...setup?.margins }
  const marginXml = Object.keys(marginValues).length ? leaf('pageMargins', { ...DEFAULT_MARGINS, ...marginValues }) : ''
  const setupXml = leaf('pageSetup', {
    ...setup?.attrs, ...(fitToPage ? setup?.fit : undefined),
  })
  const settings = leaf('printOptions', options) + marginXml + setupXml + header + rowBreaks + columnBreaks
  return { properties, settings }
}
