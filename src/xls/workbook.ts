import { caseKey } from '../collate.js'
import { InputError } from '../input-error.js'
import { readsAsCell } from '../syntax/tokens.js'
import type { XlsxIssue, XlsxIssueCode } from '../xlsx/types.js'
import { escapeAttr, escapeText } from '../xlsx/xml/text.js'
import { formulaText, type Defined, type LegacyLink } from '../xlsb/formula.js'
import { NS, type Context } from '../xlsb/parts.js'
import type { Payload } from '../xlsb/records.js'
import type { Xml } from '../xlsb/xml.js'
import { connectedCache, type ConnectedCache } from './connections.js'
import { LegacyEncoding } from './encoding.js'
import { ExternalLinks } from './external.js'
import { cacheSource, type CacheSource } from './pivot-cache.js'
import { parsedFormula } from './formula.js'
import { biffVersion, BUILTIN_NAMES, CONTINUE, Continued, EOF, isBof, records, substreamType, unicode } from './records.js'
import { Styles } from './styles.js'

export interface Sheet {
  name: string
  offset: number
  state: number
  kind: string
  end: number
  filter?: string
  links?: readonly LegacyLink[]
}

export class Losses {
  private readonly issues = new Map<string, XlsxIssue>()

  constructor(private readonly russian: boolean) {}

  add(code: XlsxIssueCode, description: string, russian: string, sheet?: string, cell?: string): void {
    const key = JSON.stringify([code, description, sheet])
    let issue = this.issues.get(key)
    if (!issue) {
      issue = { code, message: this.russian ? russian : description, ...(sheet === undefined ? {} : { sheet }) }
      this.issues.set(key, issue)
    }
    if (cell !== undefined) (issue.cells ??= []).push(cell)
  }

  list(): XlsxIssue[] { return [...this.issues.values()] }
}

export interface Workbook {
  bytes: Uint8Array
  encoding?: LegacyEncoding
  styles: Styles
  sheets: Sheet[]
  context: Context
  losses: Losses
  strings: number
  stringsXml?: string
  theme?: Uint8Array[]
  themeVersion?: number
  calc: Calculation
  external: ExternalLinks
  caches: CacheSource[]
  connectionCaches: (ConnectedCache | undefined)[]
}

export class Calculation {
  private mode?: number
  private precision?: number
  private iterate?: number
  private count?: number
  private delta?: number

  take(type: number, payload: Payload): boolean {
    switch (type) {
      case 0x000c: this.count ??= payload.u16(); return true
      case 0x000d: this.mode ??= payload.i16(); return true
      case 0x000e: this.precision ??= payload.u16(); return true
      case 0x0010: this.delta ??= payload.f64(); return true
      case 0x0011: this.iterate ??= payload.u16(); return true
      default: return false
    }
  }

  write(): string {
    const mode = this.mode === 0 ? ' calcMode="manual"' : this.mode === -1 ? ' calcMode="autoNoTable"' : ''
    const precision = this.precision === 0 ? ' fullPrecision="0"' : ''
    const iteration = this.iterate ? ' iterate="1"' : ''
    const count = this.iterate && this.count !== undefined && this.count !== 100 ? ` iterateCount="${this.count}"` : ''
    const delta = this.iterate && this.delta !== undefined && this.delta !== 0.001 ? ` iterateDelta="${this.delta}"` : ''
    const attrs = mode + precision + iteration + count + delta
    return attrs ? `<calcPr${attrs}/>` : ''
  }
}

function sheet(payload: Payload, encoding?: LegacyEncoding): Sheet {
  const offset = payload.u32()
  const state = payload.u8() & 3
  const type = payload.u8()
  const kind = type === 0 ? 'worksheet' : type === 2 ? 'chartsheet' : 'macrosheet'
  const length = payload.u8()
  const name = !length ? '' : encoding ? encoding.text(payload, length) : unicode(payload, length)
  return { offset, end: 0, state, kind, name }
}

interface BinaryName extends Defined {
  parameter: boolean
  legacyScope?: string
}

const PARAMETER_NAME = /^_xlpm\./i

function readName(payload: Payload, encoding?: LegacyEncoding, links?: readonly LegacyLink[]): BinaryName {
  const biff2 = encoding?.version === 2
  const early = encoding !== undefined && encoding.version < 5
  const flags = biff2 ? payload.u8() : payload.u16()
  payload.skip(biff2 ? 2 : 1)
  const size = payload.u8()
  const length = biff2 ? payload.u8() : payload.u16()
  let scope = -1
  let legacyScope: string | undefined
  let suffix = biff2 ? 1 : 0
  if (!early) {
    const link = payload.u16()
    scope = payload.u16() - 1
    legacyScope = links?.[link - 1]?.sheet
    suffix = payload.u8() + payload.u8() + payload.u8() + payload.u8()
  }
  const raw = encoding ? encoding.text(payload, size) : unicode(payload, size)
  const formula = encoding ? payload.section(payload.remaining - suffix) : payload
  const name = flags & 0x20 ? '_xlnm.' + (BUILTIN_NAMES[raw.charCodeAt(0)] ?? raw) : raw
  return {
    name, scope, hidden: Boolean(flags & 1), empty: !length, future: /^_xlfn\./i.test(name),
    ref: parsedFormula(formula, length, encoding, links), legacyScope,
    parameter: PARAMETER_NAME.test(name) && (flags & 0x19) === 0x19,
  }
}

function renameReservedNames(names: BinaryName[]): void {
  const reserved = names.filter(({ name, parameter }) => !parameter && (PARAMETER_NAME.test(name) || readsAsCell(name)))
  if (!reserved.length) return
  const key = (name: string, scope: number) => `${scope}:${caseKey(name)}`
  const taken = new Set(names.map(({ name, scope }) => key(name, scope)))
  for (const entry of reserved) {
    // Excel prefixes names that collide with the expanded grid or with its parameter notation.
    do { entry.name = '_' + entry.name } while (taken.has(key(entry.name, entry.scope)))
    taken.add(key(entry.name, entry.scope))
  }
}

function stringTable(source: Continued, style: Styles, xml: Xml, losses: Losses): number {
  source.u32() // Total references to strings; some writers leave this counter stale.
  const unique = source.u32()
  xml.add('<sst>')
  let read = 0
  for (; read < unique && source.hasData(); read++) {
    const length = source.u16()
    const flags = source.u8()
    const count = flags & 8 ? source.u16() : 0
    const phonetic = flags & 4 ? source.u32() : 0
    const value = source.text(length, Boolean(flags & 1))
    const text = (value: string) => `<t xml:space="preserve">${escapeText(value)}</t>`
    xml.add('<si>')
    if (count) {
      let start = 0
      let font: number | undefined
      for (let j = 0; j < count; j++) {
        const end = source.u16()
        const nextFont = source.u16()
        if (end < start || end > value.length) source.fail('rich string run outside its text')
        if (end > start) {
          const part = text(value.slice(start, end))
          xml.add(font === undefined ? `<r>${part}</r>` : style.run(font, part))
        }
        start = end
        font = nextFont
      }
      const tail = text(value.slice(start))
      xml.add(font === undefined ? tail : style.run(font, tail))
    } else xml.add(text(value))
    if (phonetic) {
      source.skip(phonetic)
      losses.add('richText', 'phonetic annotations in shared strings were left out', 'Фонетические пометки общих строк пропущены')
    }
    xml.add('</si>')
  }
  xml.add('</sst>')
  if (read !== unique) losses.add('part', 'the shared string count differs from the strings stored', 'Число общих строк отличается от сохранённых в файле')
  return read
}

export function globals(bytes: Uint8Array, xml: Xml, stringsXml: () => Xml, losses: Losses): Workbook {
  const source = records(bytes, 0, bytes.length)
  const first = source.next().value
  if (!first || !isBof(first.type)) throw new Error('xls: expected Workbook BOF')
  const declared = first.payload.u16()
  const version = biffVersion(first.type, declared)
  if (version === undefined) throw new InputError('data', 'expected an Excel BIFF workbook')
  const kind = substreamType(declared, first.payload.u16())
  const standalone = kind === 0x0010
  if (!standalone && (kind !== 0x0005 || version < 5)) first.payload.fail('expected workbook globals or a standalone worksheet')
  const encoding = version === 8 ? undefined : new LegacyEncoding(version)
  const styles = new Styles()
  const sheets: Sheet[] = standalone ? [{ name: 'Sheet1', offset: 0, end: 0, state: 0, kind: 'worksheet' }] : []
  const names: BinaryName[] = []
  const context: Context = {
    fonts: [], strings: 0, dxfs14: [], sheets: [], sheetsByRel: new Map(), names, xti: [], linkSources: [],
    external: new Map(), metadata: { cell: 0, value: 0 }, tables: new Map(),
  }
  const calc = new Calculation()
  const external = new ExternalLinks(context)
  const links = encoding?.version === 5 ? external.legacy(encoding) : undefined
  const book: Workbook = { bytes, encoding, styles, sheets, context, losses, strings: 0, calc, external, caches: [], connectionCaches: [] }
  let globalsEnd: number | undefined
  let active: number | undefined
  let date1904 = false
  let protectedBook = false
  let held = source.next().value
  const nextContinue = () => {
    held ??= source.next().value
    if (held?.type !== CONTINUE) return undefined
    const part = held.payload
    held = undefined
    return part
  }
  let embedded = 0
  let extendedXf: number | undefined
  while (held) {
    const { type, payload, offset } = held
    held = undefined
    if (type === EOF) {
      if (!embedded) {
        globalsEnd = offset + 4
        break
      }
      embedded--
      held = source.next().value
      continue
    }
    if (isBof(type)) {
      if (!standalone) payload.fail('nested BOF in workbook globals')
      embedded++
    }
    if (embedded) {
      held = source.next().value
      continue
    }
    switch (type) {
      case 0x0001: case 0x0002: case 0x0003: case 0x0004: case 0x0005: case 0x0006:
        if (encoding?.version === 2) {
          payload.skip(4)
          styles.biff2Index(payload, extendedXf)
          extendedXf = undefined
        }
        break
      case 0x0044: if (encoding?.version === 2) extendedXf = payload.u16(); break
      case 0x0020: if (encoding?.version === 2) styles.biff2ColumnStyles(payload); break
      case 0x002f: throw new InputError('data', 'password-protected .xls workbook')
      case 0x0042: if (encoding) encoding.set(payload.u16()); break
      case 0x0085: sheets.push(sheet(payload, encoding)); break
      case 0x0022: date1904 = Boolean(payload.u16()); break
      case 0x0012: case 0x0019: if (!standalone) protectedBook ||= Boolean(payload.u16()); break
      case 0x003d: if (!standalone) { payload.skip(10); active ??= payload.u16() } break
      case 0x0018: case 0x0218: names.push(readName(payload, encoding, links?.references)); break
      case 0x00fc: {
        const strings = stringsXml()
        book.strings = stringTable(new Continued(payload, nextContinue), styles, strings, losses)
        book.stringsXml = strings.text()
        break
      }
      case 0x0896: {
        payload.skip(12)
        book.themeVersion = payload.u32()
        if (payload.remaining) book.theme = [payload.bytes()]
        break
      }
      case 0x087f:
        if (book.theme && payload.u16() === 0x0896) { payload.skip(10); book.theme.push(payload.bytes()) }
        break
      case 0x00d5: book.caches.push({ id: payload.u16(), type: 0, areas: [] }); break
      case 0x0876: book.connectionCaches.push(connectedCache(payload)); break
      case 0x00d3: losses.add('macros', 'the workbook contains a VBA project', 'Книга содержит проект VBA'); break
      default: {
        const cache = book.caches.at(-1)
        if (styles.take(type, payload, encoding)) break
        if (links ? links.take(type, payload) : !encoding && external.take(type, payload, nextContinue)) break
        if (!(cache && cacheSource(cache, type, payload))) calc.take(type, payload)
      }
    }
    held ??= source.next().value
  }
  if (globalsEnd === undefined) throw new Error('xls: Workbook globals have no EOF')
  if (styles.shortenedFormat) losses.add('numberFormat', 'a number format string ends before its declared length', 'Строка числового формата короче указанной длины')
  if (standalone) sheets[0].end = globalsEnd
  const sorted = standalone ? [] : sheets.slice().sort((a, b) => a.offset - b.offset)
  for (let i = 0; i < sorted.length; i++) {
    const current = sorted[i]
    const end = sorted[i + 1]?.offset ?? bytes.length
    if (current.offset < globalsEnd || current.offset + 4 > end || end > bytes.length) throw new Error('xls: invalid sheet substream offset')
    current.end = end
  }
  for (const sheet of sheets) context.sheets.push(sheet.name)
  for (const name of names) {
    if (name.legacyScope === undefined) continue
    const key = caseKey(name.legacyScope)
    const scope = sheets.findIndex((sheet) => caseKey(sheet.name) === key)
    if (scope >= 0) name.scope = scope
  }
  renameReservedNames(names)
  xml.add(`<workbook ${NS}>`)
  if (date1904) xml.add('<workbookPr date1904="1"/>')
  if (protectedBook) xml.add('<workbookProtection/>')
  if (active) xml.add(`<bookViews><workbookView activeTab="${active}"/></bookViews>`)
  xml.add('<sheets>')
  for (let i = 0; i < sheets.length; i++) {
    const { name, state } = sheets[i]
    const hidden = state ? ` state="${state === 1 ? 'hidden' : 'veryHidden'}"` : ''
    xml.add(`<sheet name="${escapeAttr(name)}" sheetId="${i + 1}" r:id="sheet${i}"${hidden}/>`)
  }
  xml.add('</sheets><definedNames>')
  for (const name of names) {
    if (name.parameter) continue
    const formula = formulaText(name.ref, context, 0, 0, true)
    const text = formula && 'text' in formula ? formula.text : undefined
    if (!text && !name.empty && !name.future) {
      losses.add('name', `the definition of ${name.name} was left out`, `Определение имени ${name.name} пропущено`)
      continue
    }
    if (name.name === '_xlnm._FilterDatabase' && text) {
      const target = sheets[name.scope]
      if (target) target.filter = text.slice(text.lastIndexOf('!') + 1).replaceAll('$', '')
    }
    const scope = name.scope >= 0 ? ` localSheetId="${name.scope}"` : ''
    xml.add(`<definedName name="${escapeAttr(name.name)}"${scope}${name.hidden ? ' hidden="1"' : ''}>${escapeText(text ?? '')}</definedName>`)
  }
  xml.add('</definedNames>')
  // The sheet substreams may carry calculation settings, which the reader appends after finding the first sheet.
  return book
}
