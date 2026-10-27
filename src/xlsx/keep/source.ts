/**
 * The original file a book is written over: its archive, content types, workbook part with its relationships and
 * `<sheet>` entries, and each worksheet the reader read as a sheet of the book (in the book's order) with its part,
 * relationships and XML.
 */
import { attr } from '../xml/scan.js'
import { readText } from '../xml/text.js'
import { openArchive, type Archive } from '../zip/read.js'
import { relsOf, type Rel } from '../read/package.js'
import type { XlsxLimits } from '../types.js'
import { attrsOf, childrenOf, ContentTypes, rootOf, type Placed } from './xml.js'

/** A `<sheet>` of the workbook: where it lies in the workbook part, its attributes, and its part's kind and path. */
export interface SourceDecl {
  at: Placed
  attrs: Record<string, string>
  name: string
  kind?: string
  part?: string
  /** The sheet of the book it was read as, by index; none for a chart, dialog or macro sheet, or a worksheet without its part. */
  sheet?: number
}

export interface SourceSheet {
  /** Its `<sheet>` among the workbook's. */
  decl: number
  part: string
  rels: Rel[]
  /** The text of its relationships part, none without one. */
  relsXml?: string
  xml: string
}

export interface Source {
  bytes: Uint8Array
  archive: Archive
  /** The text of `[Content_Types].xml`. */
  types: string
  wbPath: string
  wb: string
  wbRels: Rel[]
  wbRelsXml?: string
  decls: SourceDecl[]
  /** The worksheets read as the book's sheets, in its order. */
  sheets: SourceSheet[]
  /** The parts of the links to other workbooks, by `<externalReference>` in order (the links' numbers); none where the part is missing. */
  links: ({ part: string; xml: string } | undefined)[]
}

/** The `.rels` part of a part. */
export const relsPath = (part: string) => {
  const dir = part.slice(0, part.lastIndexOf('/') + 1)
  return `${dir}_rels/${part.slice(dir.length)}.rels`
}

export async function openSource(bytes: Uint8Array, limits?: XlsxLimits): Promise<Source> {
  const archive = openArchive(bytes, limits)
  const types = (await archive.text('[Content_Types].xml')) ?? ''
  const root = await relsOf(archive, '')
  const wbPath = root.find((r) => r.kind === 'officeDocument' && !r.external)?.target ?? 'xl/workbook.xml'
  const wb = await archive.text(wbPath)
  if (wb === undefined) throw new Error(`xlsx: the file has no workbook part (${wbPath})`)
  const wbRels = await relsOf(archive, wbPath)
  const byId = new Map(wbRels.map((r) => [r.id, r]))
  const wbRoot = rootOf(wb)!
  const list = childrenOf(wb, wbRoot).find((c) => c.local === 'sheets')
  const names = new Set(archive.names.map((n) => n.toLowerCase()))
  const decls: SourceDecl[] = (list ? childrenOf(wb, list) : [])
    .filter((c) => c.local === 'sheet')
    .map((at) => {
      const attrs = attrsOf(wb, at)
      const r = byId.get(attr(attrs, 'id') ?? '')
      return { at, attrs, name: readText(attrs.name ?? ''), kind: r?.kind, part: r?.external ? undefined : r?.target }
    })
  const sheets: SourceSheet[] = []
  for (const [i, d] of decls.entries()) {
    // the sheets the reader reads: worksheets whose part the file has
    if (d.kind !== 'worksheet' || d.part === undefined || !names.has(d.part.toLowerCase())) continue
    const xml = await archive.text(d.part)
    if (xml === undefined) continue
    d.sheet = sheets.length
    const path = relsPath(d.part)
    sheets.push({ decl: i, part: d.part, rels: await relsOf(archive, d.part), relsXml: await archive.text(path), xml })
  }
  const refs = childrenOf(wb, wbRoot).find((c) => c.local === 'externalReferences')
  const links: Source['links'] = []
  for (const c of refs ? childrenOf(wb, refs).filter((x) => x.local === 'externalReference') : []) {
    const r = byId.get(attr(attrsOf(wb, c), 'id') ?? '')
    const xml = r && !r.external ? await archive.text(r.target) : undefined
    links.push(r && xml !== undefined ? { part: r.target, xml } : undefined)
  }
  return { bytes, archive, types, wbPath, wb, wbRels, wbRelsXml: await archive.text(relsPath(wbPath)), decls, sheets, links }
}

/** The content types of the file, to be edited. */
export const contentTypes = (src: Source) => new ContentTypes(src.types)
