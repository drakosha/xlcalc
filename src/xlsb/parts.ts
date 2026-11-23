import { escapeAttr, escapeText } from '../xlsx/xml/text.js'
import type { ExternalSource } from './external.js'
import { formula, formulaText, type Defined } from './formula.js'
import { records } from './records.js'
import type { TableIndex } from './table.js'
import type { Xml } from './xml.js'

export const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

export interface Part {
  bytes: Uint8Array
  path: string
  xml: Xml
  context: Context
  pending(): void
}

export interface Context {
  fonts: string[]
  date1904?: boolean
  strings: number
  styles?: number
  dxfs14: string[]
  sheets: string[]
  names: Defined[]
  xti: [book: number, first: number, last: number][]
  linkSources: { number?: number; sheets?: string[]; names?: string[]; scopes?: number[] }[]
  external: ReadonlyMap<string, ExternalSource>
  metadata: { cell: number; value: number }
  tables: Map<number, TableIndex>
  pivotNames?: string[]
}

export function workbook({ bytes, path, xml, pending, context }: Part): void {
  xml.add(`<workbook ${NS}>`)
  const links: string[] = []
  let sheetsOpen = false
  for (const { type, payload } of records(bytes, path)) {
    switch (type) {
      case 156: { // BrtBundleSh
        const state = payload.u32()
        const id = payload.u32()
        const rel = payload.nullableWide()
        const name = payload.wide()
        context.sheets.push(name)
        payload.done()
        if (state > 2) payload.fail(`invalid sheet state ${state}`)
        const hidden = state ? ` state="${state === 1 ? 'hidden' : 'veryHidden'}"` : ''
        if (!sheetsOpen) {
          xml.add('<sheets>')
          sheetsOpen = true
        }
        xml.add(`<sheet name="${escapeAttr(name)}" sheetId="${id}"${rel === undefined ? '' : ` r:id="${escapeAttr(rel)}"`}${hidden}/>`)
        break
      }
      case 144:
        if (sheetsOpen) xml.add('</sheets>')
        sheetsOpen = false
        break
      case 153: { // BrtWbProp
        const flags = payload.u32()
        payload.u32()
        if (payload.remaining) payload.wide()
        payload.done()
        context.date1904 = Boolean(flags & 1)
        if (context.date1904) xml.add('<workbookPr date1904="1"/>')
        break
      }
      case 158: { // BrtBookView
        payload.skip(20)
        const firstTab = payload.u32()
        const activeTab = payload.u32()
        payload.u8()
        payload.done()
        xml.add(`<bookViews><workbookView${firstTab ? ` firstSheet="${firstTab}"` : ''}${activeTab ? ` activeTab="${activeTab}"` : ''}/></bookViews>`)
        break
      }
      case 157: { // BrtCalcProp
        payload.u32()
        const mode = payload.u32()
        const count = payload.u32()
        const delta = payload.f64()
        payload.u32()
        const flags = payload.u16()
        payload.done()
        const calcMode = mode === 0 ? ' calcMode="manual"' : mode === 2 ? ' calcMode="autoNoTable"' : ''
        const fullPrecision = flags & 8 ? '' : ' fullPrecision="0"'
        const iteration = flags & 4 ? ` iterate="1" iterateCount="${count}" iterateDelta="${delta}"` : ''
        xml.add(`<calcPr${calcMode}${fullPrecision}${iteration}/>`)
        break
      }
      case 534: case 677:
        xml.add('<workbookProtection/>')
        break
      case 39: { // BrtName
        const flags = payload.u32()
        payload.u8()
        const scope = payload.u32()
        const name = payload.wide()
        const ref = formula(payload)
        while (payload.remaining) payload.nullableWide()
        const builtin = flags & 0x20 && !name.startsWith('_xlnm.') ? '_xlnm.' : ''
        context.names.push({ name: builtin + name, scope, hidden: Boolean(flags & 1), ref })
        break
      }
      case 355: { // BrtSupBookSrc
        const id = payload.nullableWide()
        payload.done()
        if (id) links.push(id)
        const external = id ? context.external.get(id) : undefined
        context.linkSources.push({ number: links.length, ...external })
        break
      }
      case 357: case 358: case 667:
        context.linkSources.push({})
        break
      case 361: {
        const source = context.linkSources.at(-1)
        const name = payload.wide()
        payload.done()
        if (!source) return payload.fail('placeholder name without a source')
        const names = (source.names ??= [])
        names.push(name)
        break
      }
      case 362: { // BrtExternSheet
        const count = payload.u32()
        if (count * 12 !== payload.remaining) payload.fail(`invalid external sheet count ${count}`)
        for (let i = 0; i < count; i++) context.xti.push([payload.u32(), payload.i32(), payload.i32()])
        break
      }
      case 384: xml.add('<pivotCaches>'); break
      case 385: xml.add('</pivotCaches>'); break
      case 386: {
        const id = payload.u32()
        const rel = payload.nullableWide()
        payload.done()
        xml.add(`<pivotCache cacheId="${id}"${rel ? ` r:id="${escapeAttr(rel)}"` : ''}/>`)
        break
      }
    }
  }
  if (sheetsOpen) xml.add('</sheets>')
  if (links.length) {
    xml.add('<externalReferences>')
    for (const id of links) xml.add(`<externalReference r:id="${escapeAttr(id)}"/>`)
    xml.add('</externalReferences>')
  }
  if (context.names.length) {
    xml.add('<definedNames>')
    for (const name of context.names) {
      if (/^_xl(?:fn|pm|op|eta)\./i.test(name.name)) continue
      if (!name.ref.tokens.remaining && !name.ref.extra.remaining) continue
      const parsed = formulaText(name.ref, context, 0, 0, true)
      if (!parsed || !('text' in parsed)) {
        pending()
        continue
      }
      const scope = name.scope === 0xffffffff ? '' : ` localSheetId="${name.scope}"`
      xml.add(`<definedName name="${escapeAttr(name.name)}"${scope}${name.hidden ? ' hidden="1"' : ''}>${escapeText(parsed.text)}</definedName>`)
    }
    xml.add('</definedNames>')
  }
  xml.add('</workbook>')
}
