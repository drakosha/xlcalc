/** Supporting links and stored external cell values ([MS-XLS] SupBook, ExternName, XCT, CRN). */
import { quoteSheetIfNeeded } from '../a1.js'
import { caseKey } from '../collate.js'
import { formatAddress } from '../model.js'
import { quoteSheet } from '../sheet-pattern.js'
import { escapeAttr, escapeText } from '../xlsx/xml/text.js'
import { errorText } from '../xlsb/fields.js'
import type { LegacyLink } from '../xlsb/formula.js'
import type { Context } from '../xlsb/parts.js'
import type { Payload } from '../xlsb/records.js'
import type { Xml } from '../xlsb/xml.js'
import { RUSSIAN_ADDIN_NAMES } from './addin-names.js'
import type { LegacyEncoding } from './encoding.js'
import { Continued, unicode } from './records.js'

type LinkSource = Context['linkSources'][number]
interface Name {
  name: string
  scope: number
  definition: string
}
interface SavedRow {
  index: number
  cells: { at: string; type?: string; value?: string }[]
}
interface SavedSheet {
  index: number
  invalid: boolean
  rows: SavedRow[]
}
interface External {
  target?: string
  sheets: string[]
  names: Name[]
  data: Map<number, SavedSheet>
}

/** A BIFF VirtualPath, used by external book links and PivotTable sources. */
export function virtualPath(value: string): { target: string; sheet?: string } | undefined {
  if (!value || value.includes('\0')) return undefined
  let filename = value
  let prefix = ''
  let transfer = false
  if (value.charCodeAt(0) === 1) {
    const kind = value.charCodeAt(1)
    filename = value.slice(1)
    if (kind === 1) {
      const volume = value[2]
      if (!volume) return undefined
      prefix = volume === '@' ? 'file://' : `file:///${volume}:/`
      filename = value.slice(3)
    } else if (kind === 2) {
      prefix = '/'
      filename = value.slice(2)
    } else if (kind === 5) {
      filename = value.slice(3)
      transfer = true
    } else if (kind >= 6 && kind <= 8) {
      // Startup, alternate startup and library paths retain their stored relative target.
      filename = value.slice(2)
    }
  } else if (value.includes('\x03')) return undefined // DDE/OLE server and topic.
  let sheet: string | undefined
  const close = filename.lastIndexOf(']')
  if (close >= 0) {
    // DConFile may bracket just the workbook name after the final directory separator.
    const open = filename.startsWith('[') ? 0
      : Math.max(filename.lastIndexOf('\x03', close), filename.lastIndexOf('/', close), filename.lastIndexOf('\\', close)) + 1
    if (open < close && filename[open] === '[') {
      sheet = filename.slice(close + 1) || undefined
      filename = filename.slice(0, open) + filename.slice(open + 1, close)
    }
  }
  let target = prefix + filename.replaceAll('\x03', '/').replaceAll('\x04', '../').replaceAll('\\', '/')
  if (transfer) {
    // Excel for Mac also writes volume paths and a single slash after the protocol in this form.
    target = target.replace(/^([a-z]):\/*/i, 'file:///$1:/').replace(/^(https?|ftp):\/(?!\/)/i, '$1://')
  }
  return { target: target.replaceAll(' ', '%20'), sheet }
}

function formulaRef(row: number, column: number): string {
  const number = String(row + 1)
  const letters = formatAddress(row, column & 255).slice(0, -number.length)
  return (column & 0x4000 ? '' : '$') + letters + (column & 0x8000 ? '' : '$') + number
}

function definition(payload: Payload, sheets: string[]): string {
  const size = payload.u16()
  if (!size) return ''
  const tokens = payload.section(size)
  const kind = tokens.u8()
  if (kind === 0x1c) return errorText(tokens.u8())
  if (kind !== 0x3a && kind !== 0x3b) return '#REF!'
  const first = sheets[tokens.u16()]
  const last = sheets[tokens.u16()]
  if (!first || !last) return '#REF!'
  const owner = first === last ? quoteSheetIfNeeded(first) : quoteSheet(first + ':' + last)
  const top = tokens.u16()
  const bottom = kind === 0x3b ? tokens.u16() : top
  const left = tokens.u16()
  const right = kind === 0x3b ? tokens.u16() : left
  const start = formulaRef(top, left)
  return owner + '!' + (kind === 0x3a ? start : start + ':' + formulaRef(bottom, right))
}

function savedValue(payload: Payload, encoding?: LegacyEncoding): { type?: string; value?: string } {
  switch (payload.u8()) {
    case 0: payload.skip(8); return {}
    case 1: return { value: String(payload.f64()) }
    case 2: return { type: 'str', value: escapeText(encoding ? encoding.text(payload, payload.u8()) : unicode(payload)) }
    case 4: {
      const value = payload.u8()
      payload.skip(7)
      return { type: 'b', value: value ? '1' : '0' }
    }
    case 16: {
      const value = errorText(payload.u8())
      payload.skip(7)
      return { type: 'e', value }
    }
    default: return payload.fail('unknown external cached value type')
  }
}

function savedRow(payload: Payload, encoding?: LegacyEncoding): SavedRow {
  const last = payload.u8()
  const first = payload.u8()
  const index = payload.u16()
  const cells: SavedRow['cells'] = []
  for (let col = first; col <= last; col++) cells.push({ at: formatAddress(index, col), ...savedValue(payload, encoding) })
  return { index, cells }
}

function legacyDefinition(payload: Payload, links: readonly LegacyLink[], book: number): string {
  const size = payload.u16()
  if (!size) return '#REF!'
  const tokens = payload.section(size)
  const kind = tokens.u8()
  if (kind === 0x1c) return errorText(tokens.u8())
  const ptg = (kind & 31) | 32
  if (ptg !== 0x3a && ptg !== 0x3b) return '#REF!'
  const link = links[tokens.u16() - 1]
  if (link?.number !== book || link.sheet === undefined) return '#REF!'
  tokens.skip(12)
  const firstRow = tokens.u16()
  const lastRow = ptg === 0x3b ? tokens.u16() : firstRow
  const firstCol = tokens.u8()
  const lastCol = ptg === 0x3b ? tokens.u8() : firstCol
  const ref = (row: number, col: number) => formulaRef(row & 0x3fff, col | (row & 0xc000))
  const start = ref(firstRow, firstCol)
  return quoteSheet(link.sheet) + '!' + (ptg === 0x3a ? start : start + ':' + ref(lastRow, lastCol))
}

function addinName(name: string, addin: boolean): string {
  const key = name.toUpperCase()
  return addin && Object.hasOwn(RUSSIAN_ADDIN_NAMES, key) ? RUSSIAN_ADDIN_NAMES[key] ?? name : name
}

interface LegacyBook {
  number: number
  book: External
  names: Set<string>
}

/** One EXTERNCOUNT table. BIFF5 numbers these entries anew in each worksheet. */
export class LegacyLinks {
  readonly references: LegacyLink[] = []
  private reference?: LegacyLink
  private target?: LegacyBook
  private sheet?: SavedSheet
  private addin = false

  constructor(
    private readonly encoding: LegacyEncoding,
    private readonly bookAt: (target: string) => LegacyBook,
  ) {}

  take(type: number, payload: Payload): boolean {
    switch (type) {
      case 0x0017: {
        const length = payload.u8()
        const prefixed = payload.remaining > length && payload.copy().u8() === 3
        const text = this.encoding.text(payload, length + +prefixed)
        const tag = text.charCodeAt(0)
        const reference: LegacyLink = {}
        this.references.push(reference)
        this.reference = reference
        this.target = this.sheet = undefined
        this.addin = text === ':'
        if (tag === 0) reference.current = true
        else if (tag === 3) reference.sheet = text.slice(1)
        else if (!this.addin && tag !== 2 && tag !== 4) {
          const target = virtualPath(text)
          if (target) {
            const external = this.bookAt(target.target)
            const { number, book } = external
            reference.number = number
            reference.sheet = target.sheet
            this.target = external
            if (target.sheet !== undefined && !book.sheets.includes(target.sheet)) book.sheets.push(target.sheet)
          }
        }
        return true
      }
      case 0x0023: {
        const reference = this.reference
        const target = this.target
        if (!reference || (!this.addin && !target)) return true
        payload.skip(2)
        const owner = this.references[payload.u16() - 1]
        payload.skip(2)
        const name = this.encoding.text(payload, payload.u8())
        const sheet = owner?.number === reference.number ? owner?.sheet : undefined
        const names = (reference.names ??= [])
        names.push({ name: addinName(name, this.addin), sheet })
        if (target) {
          const scope = sheet === undefined ? 0 : target.book.sheets.indexOf(sheet) + 1
          const key = `${scope}:${caseKey(name)}`
          if (!target.names.has(key)) {
            const definition = legacyDefinition(payload, this.references, target.number)
            target.names.add(key)
            target.book.names.push({ name, scope, definition })
          }
        }
        return true
      }
      case 0x0059: {
        const sheet = this.reference?.sheet
        const book = this.target?.book
        if (!book || sheet === undefined) return true
        const index = book.sheets.indexOf(sheet)
        const invalid = payload.i16() < 0
        const data = book.data.get(index) ?? { index, invalid: false, rows: [] }
        data.invalid ||= invalid
        book.data.set(index, data)
        this.sheet = data
        return true
      }
      case 0x005a:
        if (this.sheet) this.sheet.rows.push(savedRow(payload, this.encoding))
        return true
      default: return false
    }
  }
}

export class ExternalLinks {
  readonly links: External[] = []
  private readonly legacyBooks = new Map<string, LegacyBook>()
  private source?: LinkSource
  private addin = false
  private book?: External
  private sheet?: SavedSheet

  constructor(private readonly context: Context) {}

  legacy(encoding: LegacyEncoding): LegacyLinks {
    return new LegacyLinks(encoding, (target) => {
      const saved = this.legacyBooks.get(target)
      if (saved) return saved
      const book: External = { target, sheets: [], names: [], data: new Map() }
      const link = { book, number: this.links.push(book), names: new Set<string>() }
      this.legacyBooks.set(target, link)
      return link
    })
  }

  take(type: number, payload: Payload, next: () => Payload | undefined): boolean {
    switch (type) {
      case 0x01ae: {
        const stream = new Continued(payload, next)
        const count = stream.u16()
        const size = stream.u16()
        const source: LinkSource = {}
        this.source = source
        this.addin = size === 0x3a01
        this.context.linkSources.push(source)
        this.book = this.sheet = undefined
        if (size === 0x0401) { source.self = true; return true }
        if (this.addin) return true
        const encoded = stream.text(size, Boolean(stream.u8() & 1))
        if (encoded === '\0') { source.self = true; return true }
        const target = virtualPath(encoded)?.target
        const sheets: string[] = []
        for (let i = 0; i < count && stream.hasData(); i++) {
          const length = stream.u16()
          sheets.push(stream.text(length, Boolean(stream.u8() & 1)))
        }
        const link = { target, sheets, names: [], data: new Map<number, SavedSheet>() }
        this.book = link
        this.links.push(link)
        source.number = this.links.length
        source.sheets = sheets
        return true
      }
      case 0x0023: {
        if (!this.source || this.source.self) return true
        payload.skip(2)
        const scope = payload.u16()
        payload.skip(2)
        const name = unicode(payload, payload.u8())
        const names = (this.source.names ??= [])
        names.push(addinName(name, this.addin))
        if (this.book?.target) {
          const scopes = (this.source.scopes ??= [])
          scopes.push(scope)
          const ref = definition(payload, this.book.sheets)
          this.book.names.push({ name, scope, definition: ref })
        }
        return true
      }
      case 0x0017: {
        const stream = new Continued(payload, next)
        const count = stream.u16()
        for (let i = 0; i < count; i++) this.context.xti.push([stream.u16(), stream.i16(), stream.i16()])
        return true
      }
      case 0x0059: {
        this.sheet = undefined
        if (!this.book?.target) return true
        const rows = payload.i16()
        const index = payload.remaining >= 2 ? payload.u16() : 0
        const sheet: SavedSheet = { index, invalid: rows < 0, rows: [] }
        this.book.data.set(index, sheet)
        this.sheet = sheet
        return true
      }
      case 0x005a: {
        if (this.sheet) this.sheet.rows.push(savedRow(payload))
        return true
      }
      default: return false
    }
  }

  write(index: number, xml: Xml): string | undefined {
    const book = this.links[index]
    if (!book.target) return undefined
    xml.add('<externalLink><externalBook r:id="target"><sheetNames>')
    for (const name of book.sheets) xml.add(`<sheetName val="${escapeAttr(name)}"/>`)
    xml.add('</sheetNames><definedNames>')
    for (const { name, scope, definition } of book.names) {
      const sheet = scope ? ` sheetId="${scope - 1}"` : ''
      xml.add(`<definedName name="${escapeAttr(name)}" refersTo="${escapeAttr(definition)}"${sheet}/>`)
    }
    xml.add('</definedNames><sheetDataSet>')
    for (const sheet of book.data.values()) {
      xml.add(`<sheetData sheetId="${sheet.index}"${sheet.invalid ? ' refreshError="1"' : ''}>`)
      for (const row of sheet.rows) {
        xml.add(`<row r="${row.index + 1}">`)
        for (const { at, type, value } of row.cells) {
          const start = `<cell r="${at}"${type ? ` t="${type}"` : ''}`
          xml.add(value === undefined ? start + '/>' : start + `><v>${value}</v></cell>`)
        }
        xml.add('</row>')
      }
      xml.add('</sheetData>')
    }
    xml.add('</sheetDataSet></externalBook></externalLink>')
    return xml.text()
  }
}
