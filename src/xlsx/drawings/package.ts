/** Locate drawing parts without loading cells, styles or any image payloads. */
import { workbookInput } from '../read/book.js'
import { partText, relsOf, type Rel } from '../read/package.js'
import { repairNames } from '../read/sheet-names.js'
import type { ReadXlsxOptions } from '../types.js'
import { attr, attributes, attrsAt, within } from '../xml/scan.js'
import { decodeXml, readText } from '../xml/text.js'
import { openArchive, type Archive } from '../zip/read.js'
import type { DrawingReadIssue } from './types.js'

type SheetPart = { name: string; part: string; kind: 'worksheet' | 'chartsheet' }

export interface DrawingPart {
  sheet: SheetPart
  kind: 'drawing' | 'vmlDrawing'
  part: string
  xml: string
  rels: Rel[]
}

interface DrawingPackage {
  archive: Archive
  workbookRels: Rel[]
  sheets: SheetPart[]
}

export async function openDrawingPackage(data: Uint8Array | ArrayBuffer, opts?: ReadXlsxOptions): Promise<DrawingPackage> {
  const { bytes, limits } = workbookInput(data, opts)
  const archive = openArchive(bytes, limits)
  const root = await relsOf(archive, '')
  const part = root.find((rel) => rel.kind === 'officeDocument' && !rel.external)?.target ?? 'xl/workbook.xml'
  const workbook = await archive.text(part)
  if (workbook === undefined) throw new Error('xlsx: the file has no workbook part (' + part + ')')
  const workbookRels = await relsOf(archive, part)
  const byId = new Map(workbookRels.map((rel) => [rel.id, rel]))
  const sheets: SheetPart[] = []
  for (const node of within(workbook, 'sheets', 'sheet')) {
    const attrs = attributes(node.attrs)
    const rel = byId.get(attr(attrs, 'id') ?? '')
    if (rel && !rel.external && (rel.kind === 'worksheet' || rel.kind === 'chartsheet')) {
      sheets.push({ name: readText(attrs.name ?? ''), part: rel.target, kind: rel.kind })
    }
  }
  const worksheets = sheets.filter((sheet) => sheet.kind === 'worksheet')
  const { names } = repairNames(worksheets.map((sheet) => sheet.name))
  worksheets.forEach((sheet, index) => { sheet.name = names[index] })
  return { archive, workbookRels, sheets }
}

export async function* drawingParts(pkg: DrawingPackage, issues: DrawingReadIssue[], includeNotes: boolean): AsyncGenerator<DrawingPart> {
  for (const sheet of pkg.sheets) {
    const xml = await pkg.archive.text(sheet.part)
    if (xml === undefined) {
      issues.push({ sheet: sheet.name, part: sheet.part, reason: 'missingPart' })
      continue
    }
    const drawing = attr(attrsAt(xml, 'drawing'), 'id')
    const legacy = includeNotes ? attr(attrsAt(xml, 'legacyDrawing'), 'id') : undefined
    if (drawing === undefined && legacy === undefined) continue

    const rels = await relsOf(pkg.archive, sheet.part)
    for (const [id, kind] of [[drawing, 'drawing'], [legacy, 'vmlDrawing']] as const) {
      if (id === undefined) continue
      const rel = rels.find((item) => item.id === id && item.kind === kind && !item.external)
      const content = rel && await pkg.archive.text(rel.target)
      if (!rel || content === undefined) {
        issues.push({ sheet: sheet.name, part: rel?.target ?? sheet.part, reason: 'missingPart' })
        continue
      }
      yield { sheet, kind, part: rel.target, xml: content, rels: kind === 'drawing' ? await relsOf(pkg.archive, rel.target) : [] }
    }
  }
}

export async function pictureTypes(archive: Archive): Promise<(part: string) => string | undefined> {
  const xml = await archive.text('[Content_Types].xml') ?? ''
  const types = (tag: string, key: string) => new Map(within(xml, 'Types', tag).map((node) => {
    const attrs = attributes(node.attrs)
    return [decodeXml(attrs[key] ?? '').toLowerCase(), attrs.ContentType]
  }))
  const defaults = types('Default', 'Extension')
  const overrides = types('Override', 'PartName')
  return (part) => overrides.get('/' + part.toLowerCase()) ?? defaults.get(part.slice(part.lastIndexOf('.') + 1).toLowerCase())
}

export const drawingTheme = (pkg: DrawingPackage) => partText(pkg.archive, pkg.workbookRels, 'theme')
