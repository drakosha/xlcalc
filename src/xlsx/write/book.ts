/**
 * `writeXlsx`: a book as an xlsx file, with formulas and their values or with values only. The package holds the
 * workbook, its sheets (hidden ones with their state, the first visible one active) with their tables, notes and
 * hyperlinks, styles, shared strings, the dynamic array metadata when there are formulas, the date system and iterative
 * calculation, the links to other workbooks with the values stored of them (`links.ts`), and the book's locale as a
 * custom document property (xlsx has no place for it). A formula file asks Excel to compute the book on opening, so
 * Excel shows its own values over the ones written, whichever values it stores.
 */
import type { BookInput } from '../../contract.js'
import { activeSheetIndex } from '../view.js'
import { epochOf } from '../../dates.js'
import { SHEET_STATES } from '../../input.js'
import { FormulaError } from '../../syntax/ast.js'
import { parse } from '../../syntax/parse.js'
import { bindTables, nameProblem } from '../../syntax/check.js'
import { bindLinks } from '../../syntax/links.js'
import { localeBooleans } from '../../locale/registry.js'
import { sheetKey } from '../../syntax/print.js'
import { caseKey } from '../../collate.js'
import { toFile } from '../formulas/to-file.js'
import { Issues, type XlsxMessage } from '../issues.js'
import { LOCALE_PROPERTY } from '../property.js'
import type { WriteXlsxOptions, XlsxSnapshot, XlsxWritten } from '../types.js'
import { escapeAttr, escapeText, textElement } from '../xml/text.js'
import { utf8 } from '../zip/platform.js'
import { zip } from '../zip/write.js'
import { externalLinkXml } from './links.js'
import { nextBlock } from './parts.js'
import type { FileNames } from './formula.js'
import { writeSheet } from './sheet.js'
import { writeArgs } from './snapshot.js'
import { StyleWriter } from './styles.js'
import { tableXml } from './tables.js'
import { HEAD, MAIN } from './xml.js'
import { RICH_ERRORS } from '../read/metadata.js'
import { tableAt, tableNames, tablesOf } from '../../tables/model.js'
import { parseAddress } from '../../model.js'

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
const TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml'

const RICH = 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata'
const RICH_REL = 'http://schemas.microsoft.com/office/2017/06/relationships'
/** The rich value `errorType` of the errors Excel keeps as rich values. */
const RICH_TYPE = { SPILL: RICH_ERRORS.indexOf('#SPILL!'), CALC: RICH_ERRORS.indexOf('#CALC!') }

/**
 * The cell metadata: XLDAPR (dynamic array formulas, which `cm="1"` points at) as Excel 365 writes it, and the rich
 * values of errors (which `vm` points at, one block each).
 */
function metadata(dynamic: boolean, rich: (keyof typeof RICH_TYPE)[]): string {
  const types = [...(dynamic ? ['XLDAPR'] : []), ...(rich.length ? ['XLRICHVALUE'] : [])]
  const flags = ' minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"'
  const typeXml = types.map((t) => `<metadataType name="${t}"${flags}${t === 'XLDAPR' ? ' cellMeta="1"' : ''}/>`).join('')
  const dynamicXml = dynamic
    ? `<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>`
    : ''
  const richT = types.indexOf('XLRICHVALUE') + 1
  const richXml = rich.length
    ? `<futureMetadata name="XLRICHVALUE" count="${rich.length}">${rich.map((_, i) => `<bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="${i}"/></ext></extLst></bk>`).join('')}</futureMetadata>`
    : ''
  const cells = dynamic ? '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>' : ''
  const values = rich.length ? `<valueMetadata count="${rich.length}">${rich.map((_, i) => `<bk><rc t="${richT}" v="${i}"/></bk>`).join('')}</valueMetadata>` : ''
  return `${HEAD}<metadata xmlns="${MAIN}" xmlns:xlrd="${RICH}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="${types.length}">${typeXml}</metadataTypes>${dynamicXml}${richXml}${cells}${values}</metadata>`
}

/** A part's relationships, `rId1` on; an external one (a hyperlink's address) keeps its target as written. */
const rels = (items: [type: string, target: string, external?: boolean][]) =>
  `${HEAD}<Relationships xmlns="${PKG_REL}">${items.map(([type, target, external], i) => `<Relationship Id="rId${i + 1}" Type="${type}" Target="${external ? escapeAttr(target) : target}"${external ? ' TargetMode="External"' : ''}/>`).join('')}</Relationships>`

function definedNames(book: BookInput, issues: Issues, tables: { name: string; columns: string[] }[]): string {
  const index = new Map(book.sheets.map((s, i) => [sheetKey(s.name), i]))
  // Excel ru-RU does not load a file with a name ИСТИНА: the names the book's locale takes
  const booleans = localeBooleans(book.locale)
  const seen = new Set<string>()
  const out: string[] = []
  for (const n of book.names ?? []) {
    const scope = n.scope === undefined ? undefined : index.get(sheetKey(n.scope))
    const key = `${scope ?? ''}\n${caseKey(n.name)}`
    let text: string | undefined
    try {
      // a table's name alone is its data, which a file writes `Sales[]`
      const bound = bindTables(parse(n.ref, { union: true }), n.ref, { tables })
      // a reference to another workbook names one of the book's links and a sheet that link has
      const linked = bound.ok ? bindLinks(bound.ast, book.externalBooks, book.sheets.map((s) => s.name)) : bound
      if (linked.ok) text = toFile(linked.ast)
    } catch (e) {
      if (!(e instanceof FormulaError)) throw e
    }
    const why: XlsxMessage | undefined =
      nameProblem(n.name, booleans) !== null
        ? { key: 'nameNotTaken' }
        : n.scope !== undefined && scope === undefined
          ? { key: 'nameNoSheet' }
          : seen.has(key)
            ? { key: 'nameTwice' }
            : text === undefined
              ? { key: 'nameBadRef' }
              : undefined
    if (why) {
      issues.add('name', why, { name: n.name })
      continue
    }
    seen.add(key)
    out.push(`<definedName name="${escapeAttr(n.name)}"${scope === undefined ? '' : ` localSheetId="${scope}"`}>${escapeText(text!)}</definedName>`)
  }
  return out.length ? `<definedNames>${out.join('')}</definedNames>` : ''
}

/** Iterative calculation as `calcPr`'s attributes; what the book leaves out is Excel's default (100 rounds, a change of 0.001). */
function iterationAttrs(it: BookInput['iteration']): string {
  if (!it) return ''
  return ' iterate="1"' + (it.maxIterations !== undefined ? ` iterateCount="${it.maxIterations}"` : '') + (it.maxChange !== undefined ? ` iterateDelta="${it.maxChange}"` : '')
}

/**
 * Writes a workbook as an xlsx file: its formulas with the values from `snapshot.result` that `opts.values` requests,
 * or only its values (`opts.mode: 'values'`).
 */
export async function writeXlsx(snapshot: XlsxSnapshot, opts?: WriteXlsxOptions): Promise<XlsxWritten> {
  const { book, cells, mode, unasked, locale } = writeArgs(snapshot, opts)
  const issues = new Issues()
  const styles = new StyleWriter(book.styles)
  const strings = new Map<string, number>()
  let refs = 0
  const string = (s: string) => {
    refs++
    let i = strings.get(s)
    if (i === undefined) strings.set(s, (i = strings.size))
    return i
  }
  const rich: (keyof typeof RICH_TYPE)[] = []
  const richError = (code: keyof typeof RICH_TYPE) => (rich.includes(code) ? rich.indexOf(code) : rich.push(code) - 1) + 1
  const tables = tablesOf(book)
  const names: FileNames = { sheets: new Set(book.sheets.map((s) => sheetKey(s.name))), tables: tableNames(tables), externalBooks: book.externalBooks }
  // table parts are numbered through the book, each sheet's from rId1 of its own relationships
  let tableId = 0
  const parts = book.sheets.map((s) => (s.tables ?? []).map((t) => ({ t, id: ++tableId })))
  // a filtered sheet's criteria go on its first table with a header row, the one kind with a filter of its own
  const filterTables = book.sheets.map((s) => (s.filtered ? s.tables?.find((t) => t.headers !== false)?.name : undefined))
  // the notes' VML shape ids, in blocks of 1,024 through the book
  let noteBlock = 1
  // Excel shows the active sheet, which is never a hidden one
  const active = activeSheetIndex(book)
  const sheets = book.sheets.map((s, i) => {
    const w = writeSheet(s, i, {
      styles,
      string,
      issues,
      ...names,
      mode,
      result: cells?.[i],
      richError,
      tableAt: (addr) => {
        const { row, col } = parseAddress(addr)
        return tableAt(tables, i, row, col)?.name
      },
      totalsAt: (addr) => {
        const { row, col } = parseAddress(addr)
        const t = tableAt(tables, i, row, col)
        return !!t?.totals && row === t.r2
      },
      tableParts: parts[i].map((_, k) => `rId${k + 1}`),
      epoch: epochOf(book.dateSystem),
      ...(filterTables[i] ? { filterTable: filterTables[i] } : {}),
      noteBlock,
      active: i === active,
    })
    if (w.notes) noteBlock = nextBlock(noteBlock, w.notes.count)
    return w
  })
  const dynamic = sheets.some((s) => s.formulas)
  if (dynamic && unasked) issues.add('noCachedValues', { key: 'noResult' })
  const files: { name: string; text: string }[] = []
  const contentTypes: [string, string][] = [['/xl/workbook.xml', `${TYPE}.sheet.main+xml`]]
  const wbRels: [string, string][] = []
  let vml = false
  sheets.forEach((s, i) => {
    files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, text: s.xml })
    contentTypes.push([`/xl/worksheets/sheet${i + 1}.xml`, `${TYPE}.worksheet+xml`])
    wbRels.push([`${REL}/worksheet`, `worksheets/sheet${i + 1}.xml`])
    for (const { t, id } of parts[i]) {
      files.push({ name: `xl/tables/table${id}.xml`, text: tableXml(t, book.sheets[i], id, names, issues, t.name === filterTables[i]) })
      contentTypes.push([`/xl/tables/table${id}.xml`, `${TYPE}.table+xml`])
    }
    if (s.notes) {
      files.push({ name: `xl/comments${i + 1}.xml`, text: s.notes.comments }, { name: `xl/drawings/vmlDrawing${i + 1}.vml`, text: s.notes.vml })
      contentTypes.push([`/xl/comments${i + 1}.xml`, `${TYPE}.comments+xml`])
      vml = true
    }
    const sheetRels: [string, string, boolean?][] = [...parts[i].map(({ id }): [string, string] => [`${REL}/table`, `../tables/table${id}.xml`]), ...s.rels.map((r): [string, string, boolean] => [`${REL}/${r.type}`, r.target, r.type === 'hyperlink'])]
    if (sheetRels.length) files.push({ name: `xl/worksheets/_rels/sheet${i + 1}.xml.rels`, text: rels(sheetRels) })
  })
  files.push({ name: 'xl/styles.xml', text: styles.xml() })
  contentTypes.push(['/xl/styles.xml', `${TYPE}.styles+xml`])
  wbRels.push([`${REL}/styles`, 'styles.xml'])
  if (strings.size) {
    const items = [...strings.keys()].map((s) => `<si>${textElement(s)}</si>`)
    files.push({ name: 'xl/sharedStrings.xml', text: `${HEAD}<sst xmlns="${MAIN}" count="${refs}" uniqueCount="${strings.size}">${items.join('')}</sst>` })
    contentTypes.push(['/xl/sharedStrings.xml', `${TYPE}.sharedStrings+xml`])
    wbRels.push([`${REL}/sharedStrings`, 'sharedStrings.xml'])
  }
  if (dynamic || rich.length) {
    files.push({ name: 'xl/metadata.xml', text: metadata(dynamic, rich) })
    contentTypes.push(['/xl/metadata.xml', `${TYPE}.sheetMetadata+xml`])
    wbRels.push([`${REL}/sheetMetadata`, 'metadata.xml'])
  }
  if (rich.length) {
    const values = rich.map((code) => `<rv s="0"><v>${RICH_TYPE[code]}</v></rv>`).join('')
    files.push({ name: 'xl/richData/rdrichvalue.xml', text: `${HEAD}<rvData xmlns="${RICH}" count="${rich.length}">${values}</rvData>` })
    files.push({ name: 'xl/richData/rdrichvaluestructure.xml', text: `${HEAD}<rvStructures xmlns="${RICH}" count="1"><s t="_error"><k n="errorType" t="i"/></s></rvStructures>` })
    contentTypes.push(['/xl/richData/rdrichvalue.xml', 'application/vnd.ms-excel.rdrichvalue+xml'], ['/xl/richData/rdrichvaluestructure.xml', 'application/vnd.ms-excel.rdrichvaluestructure+xml'])
    wbRels.push([`${RICH_REL}/rdRichValue`, 'richData/rdrichvalue.xml'], [`${RICH_REL}/rdRichValueStructure`, 'richData/rdrichvaluestructure.xml'])
  }
  // the links to other workbooks, `[1]` the first, each pointing at its book through a relationship of its own
  const linkRefs = (book.externalBooks ?? []).map((b, i) => {
    const part = `externalLinks/externalLink${i + 1}.xml`
    files.push({ name: `xl/${part}`, text: externalLinkXml(b, i + 1, issues) }, { name: `xl/externalLinks/_rels/externalLink${i + 1}.xml.rels`, text: rels([[`${REL}/externalLinkPath`, b.target, true]]) })
    contentTypes.push([`/xl/${part}`, `${TYPE}.externalLink+xml`])
    return `<externalReference r:id="rId${wbRels.push([`${REL}/externalLink`, part])}"/>`
  })
  const sheetList = book.sheets.map((s, i) => `<sheet name="${escapeAttr(s.name)}" sheetId="${i + 1}"${s.hidden ? ` state="${SHEET_STATES.get(s.hidden)}"` : ''} r:id="rId${i + 1}"/>`).join('')
  // always there: without it Excel keeps the calculation settings it has from the books opened before (iteration on)
  const calc = `<calcPr${dynamic ? ' fullCalcOnLoad="1"' : ''}${iterationAttrs(book.iteration)}/>`
  const pr = book.dateSystem === 1904 ? '<workbookPr date1904="1"/>' : ''
  files.push({ name: 'xl/workbook.xml', text: `${HEAD}<workbook xmlns="${MAIN}" xmlns:r="${REL}">${pr}<bookViews><workbookView${book.activeSheet !== undefined || active ? ` activeTab="${active}"` : ''}/></bookViews><sheets>${sheetList}</sheets>${linkRefs.length ? `<externalReferences>${linkRefs.join('')}</externalReferences>` : ''}${definedNames(book, issues, names.tables)}${calc}</workbook>` })
  files.push({ name: 'xl/_rels/workbook.xml.rels', text: rels(wbRels) })
  const rootRels: [string, string][] = [[`${REL}/officeDocument`, 'xl/workbook.xml']]
  if (book.locale !== undefined) {
    const prop = `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="${LOCALE_PROPERTY}"><vt:lpwstr>${escapeText(book.locale)}</vt:lpwstr></property>`
    files.push({ name: 'docProps/custom.xml', text: `${HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">${prop}</Properties>` })
    contentTypes.push(['/docProps/custom.xml', 'application/vnd.openxmlformats-officedocument.custom-properties+xml'])
    rootRels.push([`${REL}/custom-properties`, 'docProps/custom.xml'])
  }
  files.push({ name: '_rels/.rels', text: rels(rootRels) })
  const overrides = contentTypes.map(([part, type]) => `<Override PartName="${part}" ContentType="${type}"/>`).join('')
  files.unshift({
    name: '[Content_Types].xml',
    text: `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${vml ? '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' : ''}${overrides}</Types>`,
  })
  const data = await zip(files.map((f) => ({ name: f.name, data: utf8(f.text) })))
  return { data, issues: issues.list(locale) }
}
