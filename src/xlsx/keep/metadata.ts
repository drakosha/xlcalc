/**
 * The cell metadata of a file written over another: the marks the model's cells take, the file's own where it has them
 * (`cm` of a dynamic array formula, `vm` of #SPILL! and #CALC!, which Excel 365 keeps as rich values), else the next
 * free ones, planned before the sheets are written. What a written cell takes of those the file lacks is added to its
 * metadata and rich values, or to new parts as `writeXlsx` writes them.
 */
import { cellMetadata } from '../read/metadata.js'
import { DYNAMIC, DYNAMIC_BLOCK, ERROR_STRUCTURE, metadata, metadataType, RICH, richBlock, richError, richValuesXml, RICH_TYPE } from '../write/book.js'
import { HEAD } from '../write/xml.js'
import { attributes, element, elements } from '../xml/scan.js'
import { appended, childrenOf, opening, rootOf, setAttrs, type Placed } from './xml.js'

type Rich = keyof typeof RICH_TYPE
const CODES: Rich[] = ['SPILL', 'CALC']

/** A part with items added at the end of its root element, its `count` set. */
function appendRoot(xml: string, items: string, count: number): string {
  const root = rootOf(xml)!
  return xml.slice(0, root.start) + appended(xml.slice(root.start, root.end), items, count) + xml.slice(root.end)
}

/** The marks written cells take, and whether the file has each. */
export interface MetaPlan {
  dynamic: { mark: number; has: boolean }
  rich: Record<Rich, { vm: number; has: boolean }>
}

export function planMeta(xml: string | undefined, values: string | undefined, structures: string | undefined): MetaPlan {
  const meta = cellMetadata(xml, values, structures)
  const count = (local: string) => elements(element(xml ?? '', local)?.body ?? '', 'bk').length
  let mark = 0
  for (let cm = 1; cm <= count('cellMetadata') && !mark; cm++) if (meta.dynamic(cm)) mark = cm
  const found = new Map<string, number>()
  for (let vm = count('valueMetadata'); vm >= 1; vm--) {
    const e = meta.error(vm)
    if (e) found.set(e, vm)
  }
  let next = count('valueMetadata')
  const rich = Object.fromEntries(CODES.map((c) => [c, found.has(`#${c}!`) ? { vm: found.get(`#${c}!`)!, has: true } : { vm: ++next, has: false }])) as MetaPlan['rich']
  return { dynamic: mark ? { mark, has: true } : { mark: count('cellMetadata') + 1, has: false }, rich }
}

/** The parts with what the plan adds: the metadata, and the rich values and their structures when errors are added. */
export interface MetaAdded {
  metadata: string
  values?: string
  structures?: string
}

/** What the written cells take that the file lacks, added: the dynamic arrays' mark (`dynamic`), the errors' (`rich`). */
export function addMeta(plan: MetaPlan, need: { dynamic: boolean; rich: boolean }, xml: string | undefined, values: string | undefined, structures: string | undefined): MetaAdded {
  const missing = CODES.filter((c) => !plan.rich[c].has)
  if (!xml) {
    if (values || structures) throw new Error('xlsx/keep: the file has rich values without cell metadata')
    const rich = need.rich ? missing : []
    return { metadata: metadata(need.dynamic, rich), ...(rich.length ? { values: richValuesXml(rich), structures: `${HEAD}<rvStructures xmlns="${RICH}" count="1">${ERROR_STRUCTURE}</rvStructures>` } : {}) }
  }
  const root = rootOf(xml)!
  const items: { local: string; name?: string; text: string }[] = childrenOf(xml, root).map((c) => ({ local: c.local, name: attributes(xml.slice(c.start + c.qname.length + 1, c.open)).name, text: xml.slice(c.start, c.end) }))
  const kids = (text: string, local: string) => {
    const p = rootOf(text)!
    return childrenOf(text, p).filter((c: Placed) => c.local === local).length
  }
  const typesAt = items.findIndex((i) => i.local === 'metadataTypes')
  const types = typesAt < 0 ? [] : elements(items[typesAt].text, 'metadataType').map((e) => attributes(e.attrs).name)
  const typeOf = (t: 'XLDAPR' | 'XLRICHVALUE') => {
    let i = types.indexOf(t)
    if (i < 0) {
      i = types.push(t) - 1
      const text = typesAt < 0 ? `<metadataTypes count="0"></metadataTypes>` : items[typesAt].text
      const at = typesAt < 0 ? 0 : typesAt
      items.splice(at, typesAt < 0 ? 0 : 1, { local: 'metadataTypes', name: undefined, text: appended(text, metadataType(t), types.length) })
    }
    return i + 1
  }
  /** A future metadata block of a name added: its index among that name's blocks. */
  const future = (name: string, block: string) => {
    const i = items.findIndex((x) => x.local === 'futureMetadata' && x.name === name)
    if (i >= 0) {
      const n = kids(items[i].text, 'bk')
      items[i].text = appended(items[i].text, block, n + 1)
      return n
    }
    const last = Math.max(...['metadataTypes', 'metadataStrings', 'mdxMetadata', 'futureMetadata'].map((l) => items.map((x) => x.local).lastIndexOf(l)))
    items.splice(last + 1, 0, { local: 'futureMetadata', name, text: `<futureMetadata name="${name}" count="1">${block}</futureMetadata>` })
    return 0
  }
  /** A block of cell or value metadata added. */
  const block = (local: 'cellMetadata' | 'valueMetadata', bk: string) => {
    const i = items.findIndex((x) => x.local === local)
    if (i >= 0) {
      items[i].text = appended(items[i].text, bk, kids(items[i].text, 'bk') + 1)
      return
    }
    // value metadata after the cell metadata, both after the future metadata
    const before = local === 'cellMetadata' ? ['metadataTypes', 'metadataStrings', 'mdxMetadata', 'futureMetadata'] : ['metadataTypes', 'metadataStrings', 'mdxMetadata', 'futureMetadata', 'cellMetadata']
    const last = Math.max(...before.map((l) => items.map((x) => x.local).lastIndexOf(l)))
    items.splice(last + 1, 0, { local, name: undefined, text: `<${local} count="1">${bk}</${local}>` })
  }
  if (need.dynamic && !plan.dynamic.has) {
    const t = typeOf('XLDAPR')
    block('cellMetadata', `<bk><rc t="${t}" v="${future('XLDAPR', DYNAMIC_BLOCK)}"/></bk>`)
  }
  const out: MetaAdded = { metadata: '' }
  if (need.rich && missing.length) {
    const t = typeOf('XLRICHVALUE')
    const s = structures ? kids(structures, 's') : 0
    out.structures = structures ? appendRoot(structures, ERROR_STRUCTURE, s + 1) : `${HEAD}<rvStructures xmlns="${RICH}" count="1">${ERROR_STRUCTURE}</rvStructures>`
    const n = values ? kids(values, 'rv') : 0
    const rvs = missing.map((c) => richError(c, s)).join('')
    out.values = values ? appendRoot(values, rvs, n + missing.length) : `${HEAD}<rvData xmlns="${RICH}" count="${missing.length}">${rvs}</rvData>`
    missing.forEach((_, k) => block('valueMetadata', `<bk><rc t="${t}" v="${future('XLRICHVALUE', richBlock(n + k))}"/></bk>`))
  }
  // the prefixes the blocks are written with, declared on the root where the file does not
  const tag = xml.slice(root.start, root.open)
  const a = attributes(tag.slice(root.qname.length + 1))
  const head = opening(setAttrs(tag, { ...(a['xmlns:xda'] === undefined ? { 'xmlns:xda': DYNAMIC } : {}), ...(a['xmlns:xlrd'] === undefined ? { 'xmlns:xlrd': RICH } : {}) }))
  out.metadata = xml.slice(0, root.start) + head + items.map((x) => x.text).join('') + `</${root.qname}>`
  return out
}
