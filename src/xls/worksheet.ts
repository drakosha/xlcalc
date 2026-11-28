import { bookGeometry, pointsOfChars, widthOf } from '../column-width.js'
import { MAX_COLS } from '../model.js'
import { parseRange, type Rect } from '../structure/area.js'
import { protectionAttrs } from '../xlsx/protection.js'
import { escapeText } from '../xlsx/xml/text.js'
import { errorText, rkNumber } from '../xlsb/fields.js'
import { formulaText, type FormulaText } from '../xlsb/formula.js'
import { NS } from '../xlsb/parts.js'
import type { Payload } from '../xlsb/records.js'
import type { Xml } from '../xlsb/xml.js'
import { filterMarkup, type SheetFilters } from './filter.js'
import { parsedFormula } from './formula.js'
import { Hyperlinks } from './links.js'
import type { PrintMarkup } from './printing.js'
import { address, reference, CONTINUE, Continued, EOF, isBof, substream, substreamType, unicode, type RecordInfo } from './records.js'
import { validation } from './rules.js'
import { color, fullColor } from './styles.js'
import type { Sheet, Workbook } from './workbook.js'

const PANES = ['bottomRight', 'topRight', 'bottomLeft', 'topLeft']

export function* sheetRecords(book: Workbook, sheet: Sheet): Generator<RecordInfo> {
  const source = substream(book.bytes, sheet.offset, sheet.end, (payload) => {
    const declared = payload.u16()
    if (substreamType(declared, payload.u16()) !== 0x0010) payload.fail('expected a worksheet substream')
  })
  let depth = 0
  for (const record of source) {
    if (record.type === EOF) {
      if (!depth) return
      depth--
    } else if (isBof(record.type)) depth++
    else if (!depth) yield record
  }
  throw new Error(`xls: worksheet ${sheet.name} has no EOF`)
}

function rowHeader(payload: Payload, book: Workbook, height: number | undefined): [number, string] {
  const row = payload.u16()
  payload.skip(4)
  const twips = payload.u16() & 0x7fff
  payload.skip(4)
  const flags = payload.u8()
  payload.skip(1)
  const style = flags & 128 ? book.styles.index(payload.u16() & 0xfff) : undefined
  const resize = Boolean(flags & 64)
  const hidden = Boolean(flags & 32)
  const storedHeight = twips || (hidden && resize ? height : undefined)
  let attrs = storedHeight && (resize || (height !== undefined && height !== storedHeight)) ? ` ht="${storedHeight / 20}" customHeight="${+resize}"` : ''
  if (hidden) attrs += ' hidden="1"'
  if (flags & 16) attrs += ' collapsed="1"'
  if (flags & 7) attrs += ` outlineLevel="${flags & 7}"`
  if (style !== undefined) attrs += ` s="${style}" customFormat="1"`
  return [row, attrs]
}

function biff2Row(payload: Payload, book: Workbook): [number, string] {
  const row = payload.u16()
  payload.skip(4)
  const height = payload.u16()
  payload.skip(2)
  const styled = Boolean(payload.u8())
  payload.skip(2)
  let attrs = height & 0x8000 ? '' : ` ht="${height / 20}" customHeight="1"`
  if (styled) {
    payload.skip(3)
    const style = book.styles.index(payload.u16())
    if (style !== undefined) attrs += ` s="${style}" customFormat="1"`
  }
  return [row, attrs]
}

function biff2Column(payload: Payload): Column {
  const first = payload.u8()
  const last = payload.u8()
  if (first > last) payload.fail('invalid column range')
  return { first, last, width: payload.u16() / 256, flags: 2 }
}

interface Column {
  first: number
  last: number
  width: number
  style?: number
  flags: number
}

function column(payload: Payload, book: Workbook): Column {
  const first = payload.u16()
  const last = payload.u16()
  if (first > last || first > 256) payload.fail('invalid column range')
  const width = payload.u16() / 256
  const style = book.styles.index(payload.u16())
  const flags = payload.u16()
  return { first, last, width, style, flags }
}

function columnMarkup({ first, last, width, style, flags }: Column, defaultWidth: number | undefined): string {
  const extendsGrid = last === 256
  const end = extendsGrid ? MAX_COLS : Math.min(last, 255) + 1
  const customWidth = !extendsGrid && (Boolean(flags & 2) || defaultWidth !== undefined && width !== defaultWidth)
  return `<col min="${first + 1}" max="${end}" width="${width}"`
    + (style ? ` style="${style}"` : '')
    + ` hidden="${flags & 1}" customWidth="${+customWidth}" bestFit="${(flags >>> 2) & 1}"`
    + ` outlineLevel="${(flags >>> 8) & 7}" collapsed="${(flags >>> 12) & 1}"/>`
}

function individualColumns(columns: Column[]): (Column | undefined)[] {
  const local: (Column | undefined)[] = Array(256)
  for (const column of columns) {
    for (let index = column.first; index <= Math.min(255, column.last); index++) local[index] = column
  }
  return local
}

function appendColumn(ranges: Column[], index: number, width: number, flags: number, style?: number): void {
  const previous = ranges.at(-1)
  if (previous && previous.last + 1 === index && previous.width === width && previous.flags === flags && previous.style === style) {
    previous.last = index
  } else ranges.push({ first: index, last: index, width, flags, style })
}

/** In BIFF4 and BIFF5, GCW selects the global width independently of the column's style. */
function legacyWidthColumns(
  columns: Column[], standard: number | undefined, base: number, globalWidths: Uint8Array | undefined,
): Column[] {
  const local = individualColumns(columns)
  const ranges: Column[] = []
  for (let index = 0; index < 256; index++) {
    const saved = local[index]
    const fromStandard = globalWidths === undefined || Boolean(globalWidths[index >>> 3] & (1 << (index & 7)))
    const width = fromStandard && standard !== undefined ? standard : saved?.width ?? base
    const flags = ((saved?.flags ?? 0) & ~2) | (width === standard ? 0 : 2)
    appendColumn(ranges, index, width, flags, saved?.style)
  }
  return ranges
}

function biff2Columns(columns: Column[], styles: (number | undefined)[]): Column[] {
  const local = individualColumns(columns)
  const ranges: Column[] = []
  for (let index = 0; index < 256; index++) {
    const saved = local[index]
    const style = styles[index]
    if (saved || style) appendColumn(ranges, index, saved?.width ?? 0, saved?.flags ?? 0, style)
  }
  return ranges
}

function view(payload: Payload): { flags: number; topLeft: string; zoom?: number } {
  const flags = payload.u16()
  const row = payload.u16()
  const col = payload.u16()
  payload.skip(Math.min(6, payload.remaining))
  return { flags, topLeft: address(row, col, payload), zoom: payload.remaining >= 2 ? payload.u16() : undefined }
}

function pane(payload: Payload, flags: number): string {
  const x = payload.u16()
  const y = payload.u16()
  const row = payload.u16()
  const col = payload.u16()
  const active = PANES[payload.u8()]
  const state = flags & 8 ? flags & 256 ? 'frozen' : 'frozenSplit' : 'split'
  return `<pane state="${state}" topLeftCell="${address(row, col, payload)}"`
    + (active ? ` activePane="${active}"` : '') + (x ? ` xSplit="${x}"` : '') + (y ? ` ySplit="${y}"` : '') + '/>'
}

function selection(payload: Payload): string {
  const pane = PANES[payload.u8()]
  const row = payload.u16()
  const col = payload.u16()
  const active = address(row, col, payload)
  payload.skip(2)
  const count = payload.u16()
  const refs: string[] = []
  for (let i = 0; i < count; i++) refs.push(reference(payload, true))
  return `<selection${pane ? ` pane="${pane}"` : ''}${active !== 'A1' ? ` activeCell="${active}"` : ''}`
    + (refs.length === 1 && refs[0] === 'A1' ? '' : ` sqref="${refs.join(' ')}"`) + '/>'
}

function tabColor(payload: Payload): string {
  payload.skip(12)
  const size = payload.u32()
  const index = payload.u32() & 127
  if (index === 127) return ''
  if (size === 40 && (payload.u32() & 127) === index) {
    const rgb = fullColor(payload)
    if (rgb !== 'auto="1"') return `<tabColor ${rgb}/>`
  }
  return `<tabColor ${color(index)}/>`
}

// FeatHdr's EnhancedProtection bits say which actions the protected sheet allows.
const ALLOWED_ACTIONS = [
  'objects', 'scenarios', 'formatCells', 'formatColumns', 'formatRows', 'insertColumns', 'insertRows',
  'insertHyperlinks', 'deleteColumns', 'deleteRows', 'selectLockedCells', 'sort', 'autoFilter',
  'pivotTables', 'selectUnlockedCells',
] as const

function sheetProtection(actions: number, password: number): string {
  const allowed = Object.fromEntries(ALLOWED_ACTIONS.map((action, bit) => [action, Boolean(actions & (1 << bit))]))
  const attrs = protectionAttrs(allowed)
  if (password) attrs.password = password.toString(16)
  return '<sheetProtection' + Object.entries(attrs).flatMap(([name, value]) => value === null ? [] : [` ${name}="${value}"`]).join('') + '/>'
}

export function worksheet(
  book: Workbook, sheet: Sheet, xml: Xml, hasNotes: boolean,
  relate: (id: string, kind: string, path: string, external?: boolean) => void, filters: SheetFilters, conditional: Xml | undefined, print: PrintMarkup,
): Set<string> {
  let height: number | undefined
  let baseWidth: number | undefined
  let standardWidth: number | undefined
  let details: ReturnType<typeof view> | undefined
  let paneXml = ''
  let zoom: number | undefined
  const selections: string[] = []
  const columns: Column[] = []
  const earlyStyles: (number | undefined)[] = []
  let defaultColumn: Column | undefined
  let globalWidths: Uint8Array | undefined
  const rows: [number, string][] = []
  const version = book.encoding?.version
  const biff2 = version === 2
  const legacyWidths = version === 4 || version === 5
  const tail = xml.fragment()
  const validations = xml.fragment()
  const links = new Hyperlinks()
  let zeroHeight = false
  let protectedSheet = false
  let objectsProtected = false
  let scenariosProtected = false
  let enhancedProtection: number | undefined
  let protectedRanges = false
  let password = 0
  let outlines = ''
  let tab = ''
  let inFirstView = false
  let drawing = false
  for (const { type, payload } of sheetRecords(book, sheet)) {
    switch (type) {
      case 0x0025: if (biff2) height = payload.u16() & 0x7fff; break
      case 0x0020:
        if (biff2) {
          const { first, styles } = book.styles.biff2ColumnStyles(payload)
          for (const [index, style] of styles.entries()) earlyStyles[first + index] = style
        }
        break
      case 0x0024: if (biff2) columns.push(biff2Column(payload)); break
      case 0x0008: if (biff2) rows.push(biff2Row(payload, book)); break
      case 0x0225:
        zeroHeight = Boolean(payload.u16() & 2)
        height = payload.u16()
        break
      case 0x0055: baseWidth = payload.u16(); break
      case 0x0099: standardWidth = payload.u16(); break
      case 0x00ab:
        if (legacyWidths && payload.remaining >= 34 && payload.u16() === 32) globalWidths = payload.bytes(32)
        break
      case 0x007d: {
        const saved = column(payload, book)
        if (saved.last === 256) defaultColumn = saved
        columns.push(saved)
        break
      }
      case 0x0208: rows.push(rowHeader(payload, book, height)); break
      case 0x003e:
        if (!biff2) break
        inFirstView = details === undefined
        if (inFirstView) {
          let flags = 0
          for (let bit = 0; bit < 5; bit++) if (payload.u8()) flags |= 1 << bit
          const row = payload.u16()
          const col = payload.u16()
          details = { flags, topLeft: address(row, col, payload) }
        }
        break
      case 0x023e:
        inFirstView = details === undefined
        if (inFirstView) details = view(payload)
        break
      case 0x00a0: {
        if (!inFirstView) break
        const numerator = payload.u16()
        const denominator = payload.u16()
        if (denominator) zoom = Math.round(numerator * 100 / denominator)
        break
      }
      case 0x0041: if (inFirstView) paneXml = pane(payload, details!.flags); break
      case 0x001d: if (inFirstView) selections.push(selection(payload)); break
      case 0x00e5: {
        const count = payload.u16()
        tail.add('<mergeCells>')
        for (let i = 0; i < count; i++) tail.add(`<mergeCell ref="${reference(payload)}"/>`)
        tail.add('</mergeCells>')
        break
      }
      case 0x0862: tab = tabColor(payload); break
      case 0x0012: protectedSheet = Boolean(payload.u16()); break
      case 0x0063: objectsProtected = Boolean(payload.u16()); break
      case 0x00dd: scenariosProtected = Boolean(payload.u16()); break
      case 0x0867: {
        payload.skip(12)
        if (payload.u16() !== 2) break
        payload.skip(1)
        if (payload.u32() === 0xffffffff) enhancedProtection = payload.u32()
        break
      }
      case 0x0868:
        payload.skip(12)
        if (payload.u16() === 2) protectedRanges = true
        break
      case 0x0013: password = payload.u16(); break
      case 0x0081: {
        const flags = payload.u16()
        const attrs = (flags & 32 ? ' applyStyles="1"' : '') + (flags & 64 ? '' : ' summaryBelow="0"') + (flags & 128 ? '' : ' summaryRight="0"')
        if (attrs) outlines = `<outlinePr${attrs}/>`
        break
      }
      case 0x01be: {
        const rule = validation(payload, book, sheet)
        if (rule) validations.add(rule)
        else book.losses.add('dataValidation', 'a data validation rule was left out', 'Правило проверки данных пропущено', sheet.name)
        break
      }
      case 0x005d: {
        if (!book.encoding && payload.remaining >= 6 && payload.u16() === 0x15) {
          payload.skip(2)
          const kind = payload.u16()
          if (kind !== 25 && kind !== 20) drawing = true // The dropdown for an AutoFilter is part of the sheet.
        }
        break
      }
      case 0x01b8: case 0x0800:
        if (!links.take(type, payload)) book.losses.add('hyperlink', 'a hyperlink was left out', 'Гиперссылка пропущена', sheet.name)
        break
    }
  }

  xml.add(`<worksheet ${NS}>`)
  if (filters.active || outlines || tab || print.properties) {
    xml.add(`<sheetPr${filters.active ? ' filterMode="1"' : ''}>${tab}${outlines}${print.properties ?? ''}</sheetPr>`)
  }
  if (details) {
    const attrs = (details.topLeft === 'A1' ? '' : ` topLeftCell="${details.topLeft}"`)
      + (details.flags & 2 ? '' : ' showGridLines="0"') + (details.flags & 4 ? '' : ' showRowColHeaders="0"')
    zoom ??= details.zoom
    xml.add(`<sheetViews><sheetView${attrs}${zoom && zoom !== 100 ? ` zoomScale="${zoom}"` : ''}>${paneXml}`)
    for (const value of selections) xml.add(value)
    xml.add('</sheetView></sheetViews>')
  }
  const standard = standardWidth === undefined ? undefined : standardWidth / 256
  const useLegacyWidths = legacyWidths && (standard !== undefined || globalWidths !== undefined)
  const marker = useLegacyWidths ? undefined : defaultColumn
  const defaultWidth = marker && (marker.flags & 1 ? 0 : marker.width)
  const storedWidth = defaultWidth ?? standard
  const width = storedWidth === undefined ? baseWidth === undefined ? '' : ` baseColWidth="${baseWidth}"` : ` defaultColWidth="${storedWidth}"`
  const printHeight = print && height !== undefined ? ` defaultRowHeight="${height / 20}"${zeroHeight ? ' zeroHeight="1"' : ''}` : ''
  if (width || printHeight) xml.add(`<sheetFormatPr${width}${printHeight}/>`)
  const [digit] = bookGeometry(book.styles.normalFont())
  const base = widthOf(pointsOfChars(baseWidth ?? 8, digit), digit)
  const savedColumns = useLegacyWidths ? legacyWidthColumns(columns, standard, base, globalWidths)
    : earlyStyles.length ? biff2Columns(columns, earlyStyles) : columns
  if (savedColumns.length) {
    xml.add('<cols>')
    if (marker?.first) {
      // Keep the old default for gaps to the left of the range that supplies the new default.
      const before = standard ?? base
      if (before !== defaultWidth) xml.add(`<col min="1" max="${marker.first}" width="${before}" customWidth="1"/>`)
    }
    for (const saved of savedColumns) xml.add(columnMarkup(saved, defaultWidth))
    xml.add('</cols>')
  }
  const fallback = writeCells(book, sheet, xml, rows)
  xml.append(tail)
  if (conditional) xml.append(conditional)
  if (!validations.empty) {
    xml.add('<dataValidations>')
    xml.append(validations)
    xml.add('</dataValidations>')
  }
  if (filters.enabled && sheet.filter) {
    xml.add(`<autoFilter ref="${sheet.filter}">${filterMarkup(filters.columns)}${filters.sort ?? ''}</autoFilter>`)
  }
  if (filters.rangeSort) xml.add(filters.rangeSort)
  if (protectedSheet) {
    const legacyActions = (1 << 10) | (1 << 14) | Number(!objectsProtected) | (Number(!scenariosProtected) << 1)
    xml.add(sheetProtection(enhancedProtection ?? legacyActions, password))
  }
  if (protectedRanges) xml.add('<protectedRanges/>')
  links.write(xml, relate)
  if (print.settings) xml.add(print.settings)
  if (drawing) xml.add('<drawing/>')
  if (hasNotes) xml.add('<legacyDrawing r:id="noteDrawing"/>')
  xml.add('</worksheet>')
  return fallback
}

interface Cell {
  row: number
  column: number
  at: string
  style: number | undefined
  type?: string
  value?: string
  formula?: FormulaText
  followingString?: boolean
  wasFormula?: boolean
  dataTable?: boolean
}

interface Anchor {
  kind: 'array' | 'shared'
  range: string
  id: number
  formula?: string
}

function content(cell: Cell, type: string | undefined, value: string | number): void {
  cell.type = type
  cell.value = typeof value === 'number' ? String(value) : escapeText(value)
}

function formulaCell(payload: Payload, cell: Cell, book: Workbook, sheet: Sheet): void {
  const result = payload.section(8)
  // A special result has FFFF in the last two bytes. The same bytes hold an IEEE number otherwise.
  const numeric = result.copy()
  const tag = result.u8()
  result.skip(1)
  const data = result.u8()
  result.skip(3)
  const version = book.encoding?.version
  const special = result.u16() === 0xffff
  if (version !== 2 && special) {
    switch (tag) {
      case 0: cell.followingString = true; break
      case 1: content(cell, 'b', data ? 1 : 0); break
      case 2: content(cell, 'e', errorText(data)); break
      case 3: content(cell, 'str', ''); break
      default: result.fail('invalid formula result type')
    }
  } else {
    const number = numeric.f64()
    if (!Number.isFinite(number)) result.fail('non-finite formula result')
    content(cell, undefined, number)
  }
  const flags = version === 2 ? 1 : version === 3 || version === 4 ? 2 : 6
  payload.skip(flags)
  cell.wasFormula = true
  const length = version === 2 ? payload.u8() : payload.u16()
  cell.formula = formulaText(parsedFormula(payload, length, book.encoding, sheet.links), book.context, cell.row, cell.column)
}

function richString(payload: Payload, book: Workbook): string {
  const encoding = book.encoding
  const value = encoding ? encoding.text(payload, payload.u16()) : unicode(payload)
  const count = encoding ? payload.u8() : payload.u16()
  const text = (start: number, stop?: number) => `<t xml:space="preserve">${escapeText(value.slice(start, stop))}</t>`
  if (!count) return '<is>' + text(0) + '</is>'
  let xml = '<is>'
  let start = 0
  let font: number | undefined
  const run = (part: string) => font === undefined ? '<r>' + part + '</r>' : book.styles.run(font, part)
  for (let i = 0; i < count; i++) {
    const end = encoding ? payload.u8() : payload.u16()
    const nextFont = encoding ? payload.u8() : payload.u16()
    if (end < start || end > value.length) payload.fail('rich string run outside its text')
    if (end > start) xml += run(text(start, end))
    start = end
    font = nextFont
  }
  if (start < value.length) xml += run(text(start))
  return xml + '</is>'
}

const CELL_RECORDS = new Set([0x0006, 0x00fd, 0x00d6, 0x0201, 0x0203, 0x0204, 0x0205, 0x027e, 0x00bd, 0x00be])

function writeCells(book: Workbook, sheet: Sheet, xml: Xml, rows: [number, string][]): Set<string> {
  xml.add('<sheetData>')
  rows.sort((a, b) => a[0] - b[0])
  let rowIndex = 0
  let openRow: number | undefined
  let cell: Cell | undefined
  const fallback = new Set<string>()
  const anchors = new Map<string, Anchor>()
  const version = book.encoding?.version
  const biff2 = version === 2
  let extendedXf: number | undefined
  const dataTables = new Map<string, Rect>()
  const writeEmptyRows = (before: number) => {
    while (rowIndex < rows.length && rows[rowIndex][0] < before) {
      const [index, attrs] = rows[rowIndex++]
      if (attrs) xml.add(`<row r="${index + 1}"${attrs}/>`)
    }
  }
  const writeRow = (next: number) => {
    if (next === openRow) return
    if (openRow !== undefined) xml.add('</row>')
    writeEmptyRows(next)
    openRow = next
    const attrs = rows[rowIndex]?.[0] === next ? rows[rowIndex++][1] : ''
    xml.add(`<row r="${next + 1}"${attrs}>`)
  }
  const flush = () => {
    if (!cell) return
    const { row, at, style, type, value, formula, wasFormula, dataTable, followingString } = cell
    if (followingString) throw new Error(`xls: missing string result for ${sheet.name}!${at}`)
    writeRow(row)
    let xmlFormula = dataTable ? '<f t="dataTable"/>' : ''
    let arrayMember = false
    if (!dataTable && formula && 'text' in formula) xmlFormula = `<f>${escapeText(formula.text)}</f>`
    if (!dataTable && formula && 'anchor' in formula) {
      const anchor = anchors.get(formula.anchor)
      if (!anchor) throw new Error(`xls: missing formula anchor ${sheet.name}!${formula.anchor}`)
      if (anchor.formula !== undefined) {
        if (anchor.kind === 'array') {
          arrayMember = true
          if (at === formula.anchor) xmlFormula = `<f t="array" ref="${anchor.range}">${anchor.formula}</f>`
        } else {
          xmlFormula = `<f t="shared" si="${anchor.id}"` + (at === formula.anchor ? ` ref="${anchor.range}">${anchor.formula}</f>` : '/>')
        }
      }
    }
    const table = formula && 'table' in formula ? dataTables.get(formula.table) : undefined
    const tableMember = table && row >= table.r1 && row <= table.r2 && cell.column >= table.c1 && cell.column <= table.c2
    if (wasFormula && !xmlFormula && !arrayMember && !tableMember) {
      fallback.add(at)
      book.losses.add('formula', 'a BIFF formula was left out; its stored result is used', 'Формула BIFF пропущена, использован результат из файла', sheet.name, at)
      // As a formula result, the error reaches the cache even when it cannot be typed as a cell error.
      if (type === 'e' && value !== undefined) xmlFormula = `<f>${value}</f>`
    }
    const attrs = (style === undefined ? '' : ` s="${style}"`) + (type === undefined ? '' : ` t="${type}"`)
    if (value === undefined && xmlFormula === '') xml.add(`<c r="${at}"${attrs}/>`)
    else {
      const content = value === undefined ? '' : type === 'inlineStr' ? value : `<v>${value}</v>`
      xml.add(`<c r="${at}"${attrs}>${xmlFormula}${content}</c>`)
    }
    cell = undefined
  }
  const source = sheetRecords(book, sheet)
  let next = source.next().value
  while (next) {
    const { type: record, payload } = next
    next = undefined
    let type = record
    if (biff2 && (record === 1 || record >= 3 && record <= 5)) type = record | 0x0200
    // Excel accepts the old Formula identifiers with the payload of the worksheet's BIFF version.
    if (record === 0x0206 || record === 0x0406) type = 0x0006
    if (biff2 && type === 0x0044) extendedXf = payload.u16()
    else if (CELL_RECORDS.has(type) || (biff2 && type === 0x0002)) {
      flush()
      const row = payload.u16()
      const col = payload.u16()
      if (type === 0x00bd || type === 0x00be) {
        const numeric = type === 0x00bd
        const count = (payload.remaining - 2) / (numeric ? 6 : 2)
        if (count < 1 || !Number.isInteger(count)) payload.fail('invalid multiple-cell record')
        for (let i = 0; i < count; i++) {
          cell = { row, column: col + i, at: address(row, col + i, payload), style: book.styles.index(payload.u16()) }
          if (numeric) content(cell, undefined, rkNumber(payload.u32()))
          flush()
        }
        if (payload.u16() !== col + count - 1) payload.fail('multiple-cell end column does not match the payload')
      } else {
        const style = biff2 ? book.styles.biff2Index(payload, extendedXf) : book.styles.index(payload.u16())
        extendedXf = undefined
        cell = { row, column: col, at: address(row, col, payload), style }
        switch (type) {
          case 0x0002: content(cell, undefined, payload.u16()); break
          case 0x0006: formulaCell(payload, cell, book, sheet); break
          case 0x00fd: {
            const index = payload.u32()
            if (index >= book.strings) payload.fail(`shared string index ${index} is out of bounds`)
            content(cell, 's', index)
            break
          }
          case 0x00d6: cell.type = 'inlineStr'; cell.value = richString(payload, book); break
          case 0x0203: content(cell, undefined, payload.f64()); break
          case 0x0204: content(cell, 'str', book.encoding ? book.encoding.text(payload, biff2 ? payload.u8() : payload.u16()) : unicode(payload)); break
          case 0x0205: {
            const value = payload.u8()
            if (payload.u8()) content(cell, 'e', errorText(value))
            else content(cell, 'b', value ? 1 : 0)
            break
          }
          case 0x027e: content(cell, undefined, rkNumber(payload.u32())); break
        }
      }
    } else if ((type === 0x0207 || (biff2 && type === 0x0007)) && cell?.followingString && book.encoding) {
      content(cell, 'str', book.encoding.text(payload, biff2 ? payload.u8() : payload.u16()))
      cell.followingString = false
    } else if (type === 0x0207 && cell?.followingString) {
      const string = new Continued(payload, () => {
        next = source.next().value
        if (next?.type !== CONTINUE) return undefined
        const part = next.payload
        next = undefined
        return part
      })
      const length = string.u16()
      content(cell, 'str', string.text(length, Boolean(string.u8() & 1)))
      cell.followingString = false
    } else if (type === 0x0221 || type === 0x04bc || (biff2 && type === 0x0021)) {
      if (!cell?.wasFormula) return payload.fail('formula range without its first formula cell')
      const kind = type === 0x04bc ? 'shared' : 'array'
      const range = reference(payload, true)
      const at = cell.at
      const startsAtAnchor = at === range.split(':', 1)[0]
      const early = version !== undefined && version < 5
      payload.skip(kind === 'shared' ? 2 : biff2 ? 1 : early ? 2 : 6)
      const length = biff2 ? payload.u8() : payload.u16()
      const decoded = formulaText(parsedFormula(payload, length, book.encoding, sheet.links), book.context, cell.row, cell.column, kind === 'shared')
      const formula = decoded && 'text' in decoded && (kind === 'shared' || startsAtAnchor)
        ? escapeText(decoded.text) : undefined
      anchors.set(at, { kind, range, id: anchors.size, formula })
      cell.formula = { anchor: at }
    } else if (type === 0x0236 && cell?.wasFormula) {
      const table = parseRange(reference(payload, true))!
      if (table.r1 === cell.row && table.c1 === cell.column) {
        dataTables.set(cell.at, table)
        cell.dataTable = true
      }
    }
    next ??= source.next().value
  }
  flush()
  if (openRow !== undefined) xml.add('</row>')
  writeEmptyRows(Infinity)
  xml.add('</sheetData>')
  return fallback
}
