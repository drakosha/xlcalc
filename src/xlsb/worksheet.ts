import { MAX_COLS, MAX_ROWS } from '../model.js'
import { escapeAttr, escapeText } from '../xlsx/xml/text.js'
import { Conditional } from './conditional.js'
import { address, colorAttrs, errorText, range, rkNumber } from './fields.js'
import { Filters } from './filter.js'
import { formula, formulaText, type FormulaText } from './formula.js'
import { NS, type Part } from './parts.js'
import { records, type Payload } from './records.js'
import { richString } from './strings.js'
import { validation } from './validation.js'

const PANE = ['bottomRight', 'topRight', 'bottomLeft', 'topLeft']
const PROTECTION = ['objects', 'scenarios', 'formatCells', 'formatColumns', 'formatRows', 'insertColumns',
  'insertRows', 'insertHyperlinks', 'deleteColumns', 'deleteRows', 'selectLockedCells', 'sort',
  'autoFilter', 'pivotTables', 'selectUnlockedCells']

interface Cell {
  address: string
  column: number
  attrs: string
  value: string
  formula?: FormulaText
  flags: number
  dataFormula?: string
  payload: Payload
}

interface ArrayFormula {
  array: boolean
  ref: string
  id: number
  text: string
  flags: number
}

export function worksheet(part: Part): void {
  const { bytes, path, xml, context, pending } = part
  xml.add(`<worksheet ${NS}>`)
  const filters = new Filters(xml)
  const extensions = xml.fragment()
  const hyperlinks = xml.fragment()
  const conditional = new Conditional(part)
  let row = -1
  let cm = 0
  let vm = 0
  let inData = false
  let rowOpen = false
  let defaultHeight: number | undefined
  let hasPane = false
  let colsOpen = false
  let isoProtection = false
  let validationList = false
  let lastCell: Cell | undefined
  const anchors = new Map<string, ArrayFormula>()
  const flush = () => {
    if (!lastCell) return
    const { address, attrs, value, formula, flags, dataFormula, payload } = lastCell
    const calc = flags & 2 ? ' ca="1"' : ''
    let formulaXml = dataFormula ?? ''
    if (!dataFormula && formula && 'text' in formula) formulaXml = `<f${calc}>${escapeText(formula.text)}</f>`
    if (!dataFormula && formula && 'anchor' in formula) {
      const anchor = anchors.get(formula.anchor)
      if (!anchor) return payload.fail(`formula at ${address} points to missing anchor ${formula.anchor}`)
      if (anchor.array) {
        if (address === formula.anchor) formulaXml = `<f t="array" ref="${anchor.ref}"${anchor.flags & 1 ? ' aca="1"' : ''}${calc}>${escapeText(anchor.text)}</f>`
      } else {
        formulaXml = `<f t="shared" si="${anchor.id}"${address === formula.anchor ? ` ref="${anchor.ref}">${escapeText(anchor.text)}</f>` : '/>'}`
      }
    }
    xml.add(`<c r="${address}"${attrs}>${formulaXml}${value}</c>`)
    lastCell = undefined
  }
  for (const { type, payload } of records(bytes, path)) {
    switch (type) {
      case 133: xml.add('<sheetViews>'); break
      case 134: xml.add('</sheetViews>'); break
      case 137:
        hasPane = false
        xml.add(sheetView(payload))
        break
      case 138: xml.add('</sheetView>'); break
      case 151:
        hasPane = true
        xml.add(pane(payload))
        break
      case 152:
        xml.add(selection(payload, hasPane))
        break
      case 147: xml.add(sheetProperties(payload)); break
      case 148: xml.add(`<dimension ref="${range(payload)}"/>`); payload.done(); break
      case 485: {
        const format = sheetFormat(payload)
        defaultHeight = format.height
        xml.add(format.xml)
        break
      }
      case 390: xml.add('<cols>'); colsOpen = true; break
      case 391: xml.add('</cols>'); colsOpen = false; break
      case 60:
        if (!colsOpen) payload.fail('column outside column list')
        xml.add(column(payload, context.styles))
        break
      case 145: inData = true; xml.add('<sheetData>'); break
      case 146:
        flush()
        if (rowOpen) xml.add('</row>')
        rowOpen = false
        inData = false
        xml.add('</sheetData>')
        break
      case 0: {
        if (!inData) payload.fail('row outside sheet data')
        flush()
        if (rowOpen) xml.add('</row>')
        const next = payload.u32()
        if (next >= MAX_ROWS || next <= row) payload.fail(`invalid row ${next}`)
        row = next
        rowOpen = true
        xml.add(rowHeader(payload, row, context.styles, defaultHeight))
        break
      }
      case 49:
        cm = payload.u32()
        if (cm > context.metadata.cell) payload.fail(`cell metadata index ${cm} out of bounds`)
        payload.done()
        break
      case 50:
        vm = payload.u32()
        if (vm > context.metadata.value) payload.fail(`value metadata index ${vm} out of bounds`)
        payload.done()
        break
      case 1: case 2: case 3: case 4: case 5: case 6: case 7:
      case 8: case 9: case 10: case 11: case 62:
        if (!rowOpen) payload.fail('cell without a row')
        flush()
        lastCell = cell(part, payload, type, row, cm, vm)
        cm = 0
        vm = 0
        break
      case 426: case 427: {
        if (!lastCell) return payload.fail('array or shared formula without an anchor cell')
        const ref = range(payload)
        const array = type === 426
        const flags = array ? payload.u8() : 0
        const decoded = formulaText(formula(payload), context, row, lastCell.column, !array)
        payload.done()
        if (decoded && 'anchor' in decoded) payload.fail('array or shared formula references another anchor')
        if (!decoded) pending()
        const text = decoded && 'text' in decoded ? decoded.text : ''
        if (ref.split(':', 1)[0] !== lastCell.address) payload.fail(`anchor ${lastCell.address} does not start the formula range ${ref}`)
        anchors.set(lastCell.address, { array, ref, id: anchors.size, text, flags })
        lastCell.formula = { anchor: lastCell.address }
        break
      }
      case 428:
        if (!lastCell) return payload.fail('data table without an anchor cell')
        lastCell.dataFormula = dataTable(payload)
        break
      case 177: xml.add('<mergeCells>'); break
      case 178: xml.add('</mergeCells>'); break
      case 176: xml.add(`<mergeCell ref="${range(payload)}"/>`); payload.done(); break
      case 535: {
        const protectionXml = protection(payload)
        if (!isoProtection) xml.add(protectionXml)
        isoProtection = false
        break
      }
      case 678:
        xml.add(protection(payload, true))
        isoProtection = true
        break
      case 536: case 680: case 1103: case 1104: xml.add('<protectedRanges/>'); break
      case 494:
        hyperlinks.add(hyperlink(payload))
        break
      case 550: case 551: {
        const id = payload.nullableWide()
        payload.done()
        const tag = type === 550 ? 'drawing' : 'legacyDrawing'
        xml.add(`<${tag}${id ? ` r:id="${escapeAttr(id)}"` : ''}/>`)
        break
      }
      case 392: xml.add('<rowBreaks/>'); break
      case 394: xml.add('<colBreaks/>'); break
      case 479: xml.add('<headerFooter/>'); break
      case 478: xml.add(pageSetup(payload)); break
      case 562: xml.add('<picture/>'); break
      case 638: xml.add('<oleObjects/>'); break
      case 643: xml.add('<controls/>'); break
      case 681:
        // BrtDValList belongs to the next BrtDVal, after the alternate-content block ends.
        payload.skip(payload.u32() * 2)
        payload.done()
        validationList = true
        break
      case 1058:
        payload.done()
        xml.add('<sparklineGroups/>')
        break
      default:
        if (!filters.take(type, payload) && !conditional.take(type, payload)) {
          validation(part, type, payload, extensions, validationList)
          if (type === 64) validationList = false
        }
    }
  }
  if (inData) {
    flush()
    if (rowOpen) xml.add('</row>')
    xml.add('</sheetData>')
    pending()
  }
  if (!hyperlinks.empty) {
    xml.add('<hyperlinks>')
    xml.append(hyperlinks)
    xml.add('</hyperlinks>')
  }
  conditional.appendExtensions(extensions)
  if (!extensions.empty) {
    xml.add('<extLst xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">')
    xml.append(extensions)
    xml.add('</extLst>')
  }
  xml.add('</worksheet>')
}

function sheetProperties(payload: Payload): string {
  const flags = payload.u16()
  const extendedFlags = payload.u8()
  const color = colorAttrs(payload)
  payload.skip(8)
  const codeName = payload.wide()
  payload.done()
  const autoStyle = flags & 0x20 ? ' applyStyles="1"' : ''
  const summaryBelow = flags & 0x40 ? '' : ' summaryBelow="0"'
  const summaryRight = flags & 0x80 ? '' : ' summaryRight="0"'
  const outline = autoStyle || summaryBelow || summaryRight ? `<outlinePr${autoStyle}${summaryBelow}${summaryRight}/>` : ''
  const filter = extendedFlags & 1 ? ' filterMode="1"' : ''
  const tab = color === 'auto="1"' || color === 'indexed="64"' ? '' : `<tabColor ${color}/>`
  const code = codeName ? ` codeName="${escapeAttr(codeName)}"` : ''
  return `<sheetPr${filter}${code}>${tab}${outline}</sheetPr>`
}

function sheetView(payload: Payload): string {
  const flags = payload.u16()
  payload.u32()
  const row = payload.u32()
  const col = payload.u32()
  payload.skip(4)
  const zoom = payload.u16()
  payload.skip(6)
  const book = payload.u32()
  payload.done()
  const top = address(payload, row, col)
  return `<sheetView workbookViewId="${book}"${top === 'A1' ? '' : ` topLeftCell="${top}"`}${zoom === 100 ? '' : ` zoomScale="${zoom}"`}${flags & 4 ? '' : ' showGridLines="0"'}${flags & 8 ? '' : ' showRowColHeaders="0"'}>`
}

function pane(payload: Payload): string {
  const x = payload.f64()
  const y = payload.f64()
  const row = payload.u32()
  const col = payload.u32()
  const active = PANE[payload.u32()]
  const flags = payload.u8()
  payload.done()
  if (!active) payload.fail('invalid active pane')
  const state = flags & 2 ? 'frozen' : flags & 1 ? 'frozenSplit' : 'split'
  return `<pane topLeftCell="${address(payload, row, col)}" state="${state}" activePane="${active}"${x ? ` xSplit="${x}"` : ''}${y ? ` ySplit="${y}"` : ''}/>`
}

function selection(payload: Payload, hasPane: boolean): string {
  const selectedPane = PANE[payload.u32()]
  const row = payload.u32()
  const col = payload.u32()
  const index = payload.u32()
  const count = payload.u32()
  if (count > 32767 || count * 16 > payload.remaining || index >= count) payload.fail('invalid selection range table')
  const active = address(payload, row, col)
  const ranges: string[] = []
  for (let i = 0; i < count; i++) ranges.push(range(payload))
  payload.done()
  if (!selectedPane) payload.fail('invalid selection pane')
  const paneAttr = hasPane ? ` pane="${selectedPane}"` : ''
  const activeAttr = active === 'A1' ? '' : ` activeCell="${active}"`
  const rangesAttr = ranges.length === 1 && ranges[0] === 'A1' ? '' : ` sqref="${ranges.join(' ')}"`
  return paneAttr || activeAttr || rangesAttr ? `<selection${paneAttr}${activeAttr}${rangesAttr}/>` : ''
}

function sheetFormat(payload: Payload): { height: number; xml: string } {
  const width = payload.u32()
  const base = payload.u16()
  const height = payload.u16()
  const flags = payload.u16()
  const rowOutline = payload.u8()
  const colOutline = payload.u8()
  payload.done()
  const colWidth = width === 0xffffffff ? ` baseColWidth="${base}"` : ` defaultColWidth="${width / 256}"`
  const xml = `<sheetFormatPr${colWidth} defaultRowHeight="${height / 20}" customHeight="${flags & 1}" outlineLevelRow="${rowOutline}" outlineLevelCol="${colOutline}"/>`
  return { height, xml }
}

function styleAttr(payload: Payload, style: number, count: number | undefined, attr: string): string {
  if (count !== undefined && style >= count) payload.fail(`style index ${style} out of bounds`)
  return style ? ` ${attr}="${style}"` : ''
}

function column(payload: Payload, styles: number | undefined): string {
  const first = payload.u32()
  const last = payload.u32()
  const width = payload.u32()
  const style = styleAttr(payload, payload.u32(), styles, 'style')
  const flags = payload.u16()
  payload.done()
  if (first > last || last >= MAX_COLS || width > 65535) payload.fail('invalid column range or width')
  return `<col min="${first + 1}" max="${last + 1}" width="${width / 256}"${style} hidden="${flags & 1}" customWidth="${(flags >>> 1) & 1}" bestFit="${(flags >>> 2) & 1}" outlineLevel="${(flags >>> 8) & 7}" collapsed="${(flags >>> 12) & 1}"/>`
}

function rowHeader(payload: Payload, row: number, styles: number | undefined, defaultHeight: number | undefined): string {
  const index = payload.u32()
  const height = payload.u16()
  payload.u8()
  const flags = payload.u8()
  payload.u8()
  const count = payload.u32()
  if (count > 16) payload.fail(`invalid column span count ${count}`)
  payload.skip(count * 8)
  payload.done()
  const style = flags & 0x40 ? styleAttr(payload, index, styles, 's') + ' customFormat="1"' : ''
  const manual = flags & 0x20 ? ' customHeight="1"' : ''
  // Excel saves a fitted height without the manual flag when it differs from the sheet's default.
  const heightAttr = manual || (height && defaultHeight !== undefined && height !== defaultHeight) ? ` ht="${height / 20}"` : ''
  return `<row r="${row + 1}"${style}${heightAttr}${manual}${flags & 0x10 ? ' hidden="1"' : ''}${flags & 8 ? ' collapsed="1"' : ''}${flags & 7 ? ` outlineLevel="${flags & 7}"` : ''}>`
}

function cell({ context, pending }: Part, payload: Payload, type: number, row: number, cm: number, vm: number): Cell {
  const column = payload.u32()
  const addr = address(payload, row, column)
  const style = styleAttr(payload, payload.u32() & 0xffffff, context.styles, 's')
  const metadata = (cm ? ` cm="${cm}"` : '') + (vm ? ` vm="${vm}"` : '')
  let kind = ''
  let body = ''
  let value: number | string | undefined
  switch (type) {
    case 2: value = rkNumber(payload.u32()); break
    case 3: case 11: kind = 'e'; value = errorText(payload.u8()); break
    case 4: case 10: kind = 'b'; value = payload.u8(); break
    case 5: case 9: value = payload.f64(); break
    case 6: case 8: kind = 'str'; value = escapeText(payload.wide()); break
    case 7:
      kind = 's'
      value = payload.u32()
      if (value >= context.strings) payload.fail(`shared string index ${value} out of bounds`)
      break
    case 62:
      kind = 'inlineStr'
      body = `<is>${richString(payload, context.fonts, pending)}</is>`
      break
  }
  const calculated = type >= 8 && type <= 11
  const flags = calculated ? payload.u16() : 0
  const decoded = calculated ? formulaText(formula(payload), context, row, column) : undefined
  if (calculated && !decoded) pending()
  payload.done()
  if (value !== undefined) body += `<v>${value}</v>`
  return { address: addr, column, attrs: style + metadata + (kind ? ` t="${kind}"` : ''), value: body, formula: decoded, flags, payload }
}

function dataTable(payload: Payload): string {
  const ref = range(payload)
  const row1 = payload.u32()
  const col1 = payload.u32()
  const row2 = payload.u32()
  const col2 = payload.u32()
  const flags = payload.u8()
  payload.done()
  const fields = (flags & 4 ? ' del1="1"' : ` r1="${address(payload, row1, col1)}"`)
    + (flags & 2 ? ` dt2D="1"${flags & 8 ? ' del2="1"' : ` r2="${address(payload, row2, col2)}"`}` : '')
    + (flags & 1 ? ' dtr="1"' : '')
    + (flags & 0x10 ? ' ca="1"' : '')
  return `<f t="dataTable" ref="${ref}"${fields}/>`
}

function protection(payload: Payload, iso = false): string {
  const count = iso ? payload.u32() : undefined
  const password = iso ? 0 : payload.u16()
  const locked = payload.u32()
  const attrs = PROTECTION.map((prop) => ` ${prop}="${payload.u32() ? 0 : 1}"`).join('')
  let passwordAttr = password ? ` password="${password.toString(16)}"` : ''
  if (iso) {
    const hashBytes = payload.u32()
    payload.skip(hashBytes)
    payload.skip(payload.u32()) // salt
    const algorithm = payload.nullableWide()
    if (hashBytes && algorithm) passwordAttr = ` algorithmName="${escapeAttr(algorithm)}" spinCount="${count}"`
  }
  payload.done()
  return locked ? `<sheetProtection sheet="1"${passwordAttr}${attrs}/>` : ''
}

function hyperlink(payload: Payload): string {
  const ref = range(payload)
  const id = payload.nullableWide()
  const tokens = [['location', payload.wide()], ['tooltip', payload.wide()], ['display', payload.wide()]]
  payload.done()
  const attrs = tokens.map(([key, value]) => value ? ` ${key}="${escapeAttr(value)}"` : '').join('')
  return `<hyperlink ref="${ref}"${id ? ` r:id="${escapeAttr(id)}"` : ''}${attrs}/>`
}

/** The setup options that readXlsx reports, whose binary defaults Excel leaves out of XML. */
function pageSetup(payload: Payload): string {
  payload.skip(4)
  const scale = payload.u32()
  payload.skip(16)
  const width = payload.u32()
  const height = payload.u32()
  const flags = payload.u16()
  payload.nullableWide()
  payload.done()
  const orientation = (flags & 0x42) === 2 ? ' orientation="landscape"' : ''
  const sizing = (scale === 100 ? '' : ` scale="${scale}"`)
    + (width === 1 ? '' : ` fitToWidth="${width}"`) + (height === 1 ? '' : ` fitToHeight="${height}"`)
  return orientation || sizing ? `<pageSetup${orientation}${sizing}/>` : ''
}
