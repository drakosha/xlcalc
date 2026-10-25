/**
 * Files for the tests of `xlcalc/xlsx/keep`: a book written by `writeXlsx`, with parts the model does not have added as
 * Excel writes them (a drawing with a chart and a picture, a pivot cache and its table, the print area, page breaks,
 * formatted runs of text, a cell format with rotation), and the parts of a file as text to compare.
 */
import type { BookInput } from '../src/contract.js'
import { compute } from '../src/index.js'
import { writeXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { zip } from '../src/xlsx/zip/write.js'

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const DRAW = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const TYPE = 'application/vnd.openxmlformats-officedocument'

/** Rows and columns from 0: the top left and bottom right cells of a drawing's anchor. */
export type Anchor = [r1: number, c1: number, r2: number, c2: number]

export interface Extra {
  /** A chart on a sheet (by index) over the series' references, at an anchor, moving and sizing with its cells unless `editAs` says otherwise. */
  chart?: { sheet: number; series: string[]; anchor: Anchor; editAs?: string }
  /** A picture on a sheet, moving with its cells (`editAs` oneCell unless it says otherwise). */
  picture?: { sheet: number; anchor: Anchor; editAs?: string }
  /** A pivot table on a sheet at `location`, its cache's source cells or the table or name it takes them from. */
  pivot?: { sheet: number; location: string; source: { sheet: string; ref: string } | { name: string } }
  /** `_xlnm.Print_Area` of a sheet. */
  printArea?: { sheet: number; ref: string }
  /** Page breaks before rows (0-based ids as Excel writes them). */
  breaks?: { sheet: number; rows: number[] }
  /** Shared strings given formatted runs: the text and the runs' XML. */
  rich?: Record<string, string>
  /** Cell formats given a rotation of their text, by index. */
  rotated?: number[]
}

const anchorXml = (x: Anchor, inner: string, editAs?: string) =>
  `<xdr:twoCellAnchor${editAs ? ` editAs="${editAs}"` : ''}><xdr:from><xdr:col>${x[1]}</xdr:col><xdr:colOff>1000</xdr:colOff><xdr:row>${x[0]}</xdr:row><xdr:rowOff>2000</xdr:rowOff></xdr:from><xdr:to><xdr:col>${x[3]}</xdr:col><xdr:colOff>1000</xdr:colOff><xdr:row>${x[2]}</xdr:row><xdr:rowOff>2000</xdr:rowOff></xdr:to>${inner}<xdr:clientData/></xdr:twoCellAnchor>`

/** A file of `writeXlsx`'s for the book and its values, with the extra parts. */
export async function withParts(book: BookInput, extra: Extra): Promise<Uint8Array> {
  const base = openArchive((await writeXlsx({ book, result: compute(book) })).data)
  const parts = new Map<string, string>()
  for (const name of base.names) parts.set(name, (await base.text(name))!)
  const edit = (name: string, f: (x: string) => string) => parts.set(name, f(parts.get(name) ?? ''))
  const override = (part: string, type: string) => edit('[Content_Types].xml', (x) => x.replace('</Types>', `<Override PartName="/${part}" ContentType="${type}"/></Types>`))
  const sheetRel = (i: number, id: string, kind: string, target: string) => {
    const name = `xl/worksheets/_rels/sheet${i + 1}.xml.rels`
    if (!parts.has(name)) parts.set(name, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`)
    edit(name, (x) => x.replace('</Relationships>', `<Relationship Id="${id}" Type="${REL}/${kind}" Target="${target}"/></Relationships>`))
  }
  /** An element put into a sheet before the first of the elements the schema puts after it. */
  const inSheet = (i: number, xml: string, after: string[]) =>
    edit(`xl/worksheets/sheet${i + 1}.xml`, (x) => {
      const at = Math.min(...after.map((t) => x.indexOf(`<${t}`)).filter((k) => k >= 0), x.indexOf('</worksheet>'))
      return x.slice(0, at) + xml + x.slice(at)
    })
  const anchors = new Map<number, string[]>()
  if (extra.chart) {
    const { sheet, series, anchor } = extra.chart
    const ser = series.map((f, k) => `<c:ser><c:idx val="${k}"/><c:order val="${k}"/><c:val><c:numRef><c:f>${f}</c:f></c:numRef></c:val></c:ser>`).join('')
    parts.set('xl/charts/chart1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${REL}"><c:chart><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser}<c:axId val="1"/><c:axId val="2"/></c:barChart></c:plotArea></c:chart></c:chartSpace>`)
    override('xl/charts/chart1.xml', `${TYPE}.drawingml.chart+xml`)
    ;(anchors.get(sheet) ?? anchors.set(sheet, []).get(sheet)!).push(anchorXml(anchor, `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" r:id="rIdC1"/></a:graphicData></a:graphic></xdr:graphicFrame>`, extra.chart.editAs))
  }
  if (extra.picture) {
    const { sheet, anchor } = extra.picture
    parts.set('xl/media/image1.png', 'png bytes')
    edit('[Content_Types].xml', (x) => x.replace('<Default Extension="xml"', '<Default Extension="png" ContentType="image/png"/><Default Extension="xml"'))
    ;(anchors.get(sheet) ?? anchors.set(sheet, []).get(sheet)!).push(anchorXml(anchor, `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="3" name="Picture 1"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdP1"/></xdr:blipFill><xdr:spPr/></xdr:pic>`, extra.picture.editAs ?? 'oneCell'))
  }
  for (const [sheet, list] of anchors) {
    parts.set('xl/drawings/drawing1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${DRAW}" xmlns:a="${A}" xmlns:r="${REL}">${list.join('')}</xdr:wsDr>`)
    const rels = [...(extra.chart ? [`<Relationship Id="rIdC1" Type="${REL}/chart" Target="../charts/chart1.xml"/>`] : []), ...(extra.picture ? [`<Relationship Id="rIdP1" Type="${REL}/image" Target="../media/image1.png"/>`] : [])]
    parts.set('xl/drawings/_rels/drawing1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`)
    override('xl/drawings/drawing1.xml', `${TYPE}.drawing+xml`)
    sheetRel(sheet, 'rIdD1', 'drawing', '../drawings/drawing1.xml')
    inSheet(sheet, '<drawing r:id="rIdD1"/>', ['legacyDrawing', 'tableParts', 'extLst'])
  }
  if (extra.pivot) {
    const { sheet, location, source } = extra.pivot
    const src = 'name' in source ? `<worksheetSource name="${source.name}"/>` : `<worksheetSource ref="${source.ref}" sheet="${source.sheet}"/>`
    parts.set('xl/pivotCache/pivotCacheDefinition1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL}" refreshOnLoad="1"><cacheSource type="worksheet">${src}</cacheSource><cacheFields count="0"/></pivotCacheDefinition>`)
    override('xl/pivotCache/pivotCacheDefinition1.xml', `${TYPE}.spreadsheetml.pivotCacheDefinition+xml`)
    parts.set('xl/pivotTables/pivotTable1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="1"><location ref="${location}" firstHeaderRow="1" firstDataRow="1" firstDataCol="0"/></pivotTableDefinition>`)
    parts.set('xl/pivotTables/_rels/pivotTable1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>`)
    override('xl/pivotTables/pivotTable1.xml', `${TYPE}.spreadsheetml.pivotTable+xml`)
    sheetRel(sheet, 'rIdT1', 'pivotTable', '../pivotTables/pivotTable1.xml')
    edit('xl/_rels/workbook.xml.rels', (x) => x.replace('</Relationships>', `<Relationship Id="rIdPC" Type="${REL}/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/></Relationships>`))
    edit('xl/workbook.xml', (x) => x.replace(/(<calcPr[^>]*\/>)/, '$1<pivotCaches><pivotCache cacheId="1" r:id="rIdPC"/></pivotCaches>'))
  }
  if (extra.printArea) {
    const n = `<definedName name="_xlnm.Print_Area" localSheetId="${extra.printArea.sheet}">${extra.printArea.ref}</definedName>`
    edit('xl/workbook.xml', (x) => (x.includes('<definedNames>') ? x.replace('<definedNames>', `<definedNames>${n}`) : x.replace('<calcPr', `<definedNames>${n}</definedNames><calcPr`)))
  }
  if (extra.breaks) {
    const b = extra.breaks.rows.map((r) => `<brk id="${r}" max="16383" man="1"/>`).join('')
    inSheet(extra.breaks.sheet, `<rowBreaks count="${extra.breaks.rows.length}" manualBreakCount="${extra.breaks.rows.length}">${b}</rowBreaks>`, ['drawing', 'legacyDrawing', 'tableParts', 'extLst'])
  }
  for (const [text, runs] of Object.entries(extra.rich ?? {})) edit('xl/sharedStrings.xml', (x) => x.replace(`<si><t>${text}</t></si>`, `<si>${runs}</si>`))
  for (const i of extra.rotated ?? [])
    edit('xl/styles.xml', (x) => {
      let k = -1
      return x.replace(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/, (all) => all.replace(/<xf [^>]*?(\/>|>[\s\S]*?<\/xf>)/g, (xf) => (++k === i ? `${xf.replace(/\/>$|>[\s\S]*<\/xf>$/, '')} applyAlignment="1"><alignment textRotation="45"/></xf>` : xf)))
    })
  return zip([...parts].map(([name, text]) => ({ name, data: new TextEncoder().encode(text) })))
}

/** The parts of a file as text, by name. */
export async function partsOf(data: Uint8Array): Promise<Map<string, string>> {
  const a = openArchive(data)
  const out = new Map<string, string>()
  for (const name of a.names) out.set(name, (await a.text(name))!)
  return out
}

/** The names of the parts two files differ in, added and removed ones among them. */
export async function changedParts(a: Uint8Array, b: Uint8Array): Promise<string[]> {
  const [x, y] = [await partsOf(a), await partsOf(b)]
  return [...new Set([...x.keys(), ...y.keys()])].filter((n) => x.get(n) !== y.get(n)).sort()
}
