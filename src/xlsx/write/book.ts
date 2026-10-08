/**
 * `writeXlsx`: a book as an xlsx file, with formulas and their values or with values only. The package holds the
 * workbook, its sheets, styles, shared strings, the dynamic array metadata when there are formulas, the date system and
 * iterative calculation, and the book's locale as a custom document property (xlsx has no place for it). A formula
 * file asks Excel to compute the book on opening, so Excel shows its own values over the ones written.
 */
import type { BookInput, ComputeResult } from '../../contract.js'
import { epochOf } from '../../dates.js'
import { checkBook, InputError } from '../../input.js'
import { FormulaError } from '../../syntax/ast.js'
import { parse } from '../../syntax/parse.js'
import { bindTables } from '../../syntax/check.js'
import { sheetKey } from '../../syntax/print.js'
import { toFile } from '../formulas/to-file.js'
import { Issues } from '../issues.js'
import { LOCALE_PROPERTY } from '../property.js'
import type { WriteXlsxOptions, XlsxWritten } from '../types.js'
import { escapeAttr, escapeText, needsPreserve } from '../xml/text.js'
import { utf8 } from '../zip/platform.js'
import { zip } from '../zip/write.js'
import { writeSheet } from './sheet.js'
import { StyleWriter } from './styles.js'
import { tableXml } from './tables.js'
import { tableAt, tablesOf } from '../../tables/model.js'
import { parseAddress } from '../../model.js'

const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
const TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml'

const RICH = 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata'
const RICH_REL = 'http://schemas.microsoft.com/office/2017/06/relationships'
/** The rich value `errorType` of the errors Excel keeps as rich values. */
const RICH_TYPE = { SPILL: 8, CALC: 13 } as const

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

const rels = (items: [type: string, target: string][]) =>
  `${HEAD}<Relationships xmlns="${PKG_REL}">${items.map(([type, target], i) => `<Relationship Id="rId${i + 1}" Type="${type}" Target="${target}"/>`).join('')}</Relationships>`

/** A name the file holds: Excel's rules (letters of any script, `_` or `\` first; not a cell address, R1C1 or a boolean). */
const NAME = /^[\p{L}_\\][\p{L}\p{N}_.\\?]*$/u
const NOT_NAME = /^(?:[A-Z]{1,3}\d+|R\d*C\d*|R|C|TRUE|FALSE)$/i

function definedNames(book: BookInput, issues: Issues, tables: { name: string; columns: string[] }[]): string {
  const index = new Map(book.sheets.map((s, i) => [sheetKey(s.name), i]))
  const seen = new Set<string>()
  const out: string[] = []
  for (const n of book.names ?? []) {
    const scope = n.scope === undefined ? undefined : index.get(sheetKey(n.scope))
    const key = `${scope ?? ''}\n${n.name.toUpperCase()}`
    let text: string | undefined
    try {
      // a table's name alone is its data, which a file writes `Sales[]`
      const bound = bindTables(parse(n.ref, { union: true }), n.ref, { tables })
      if (bound.ok) text = toFile(bound.ast)
    } catch (e) {
      if (!(e instanceof FormulaError)) throw e
    }
    const why =
      !NAME.test(n.name) || NOT_NAME.test(n.name) || n.name.length > 255
        ? 'a name Excel does not take'
        : n.scope !== undefined && scope === undefined
          ? 'a name of a sheet the book does not have'
          : seen.has(key)
            ? 'a name given twice'
            : text === undefined
              ? 'a name whose reference does not parse or names a table or a column the book does not have'
              : undefined
    if (why) {
      issues.add('name', `${why} is left out of the file`, { name: n.name })
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

/** The result's cells by sheet of the book; the result must be the book's (same sheets, in order). */
function resultCells(book: BookInput, result: ComputeResult | undefined) {
  if (!result) return undefined
  if (result.sheets.length !== book.sheets.length || result.sheets.some((s, i) => sheetKey(s.name) !== sheetKey(book.sheets[i].name)))
    throw new InputError('opts.result', "expected the book's result: its sheets differ from the book's")
  return result.sheets.map((s) => s.cells)
}

/** Writes a book as an xlsx file: with formulas and the values of `result`, or with values only. */
export async function writeXlsx(book: BookInput, opts: WriteXlsxOptions = {}): Promise<XlsxWritten> {
  checkBook(book)
  const mode = opts.mode ?? 'formulas'
  if (mode !== 'formulas' && mode !== 'values') throw new InputError('opts.mode', `expected "formulas" or "values", got ${JSON.stringify(mode)}`)
  if (mode === 'values' && !opts.result) throw new InputError('opts.result', 'the values mode takes the values from the result: pass the result of compute')
  const cells = resultCells(book, opts.result)
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
  const sheetKeys = new Set(book.sheets.map((s) => sheetKey(s.name)))
  const rich: (keyof typeof RICH_TYPE)[] = []
  const richError = (code: keyof typeof RICH_TYPE) => (rich.includes(code) ? rich.indexOf(code) : rich.push(code) - 1) + 1
  const tables = tablesOf(book)
  const tableNames = tables.list.map((t) => ({ name: t.name, columns: t.cols }))
  // table parts are numbered through the book, each sheet's from rId1 of its own relationships
  let tableId = 0
  const parts = book.sheets.map((s) => (s.tables ?? []).map((t) => ({ t, id: ++tableId })))
  // a filtered sheet's criteria go on its first table with a header row, the one kind with a filter of its own
  const filterTables = book.sheets.map((s) => (s.filtered ? s.tables?.find((t) => t.headers !== false)?.name : undefined))
  const sheets = book.sheets.map((s, i) =>
    writeSheet(s, i, {
      styles,
      string,
      issues,
      sheets: sheetKeys,
      mode,
      result: cells?.[i],
      richError,
      tables: tableNames,
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
    }),
  )
  const dynamic = sheets.some((s) => s.formulas)
  const files: { name: string; text: string }[] = []
  const contentTypes: [string, string][] = [['/xl/workbook.xml', `${TYPE}.sheet.main+xml`]]
  const wbRels: [string, string][] = []
  sheets.forEach((s, i) => {
    files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, text: s.xml })
    contentTypes.push([`/xl/worksheets/sheet${i + 1}.xml`, `${TYPE}.worksheet+xml`])
    wbRels.push([`${REL}/worksheet`, `worksheets/sheet${i + 1}.xml`])
    if (!parts[i].length) return
    for (const { t, id } of parts[i]) {
      files.push({ name: `xl/tables/table${id}.xml`, text: tableXml(t, book.sheets[i], id, sheetKeys, tableNames, t.name === filterTables[i]) })
      contentTypes.push([`/xl/tables/table${id}.xml`, `${TYPE}.table+xml`])
    }
    files.push({ name: `xl/worksheets/_rels/sheet${i + 1}.xml.rels`, text: rels(parts[i].map(({ id }) => [`${REL}/table`, `../tables/table${id}.xml`])) })
  })
  files.push({ name: 'xl/styles.xml', text: styles.xml() })
  contentTypes.push(['/xl/styles.xml', `${TYPE}.styles+xml`])
  wbRels.push([`${REL}/styles`, 'styles.xml'])
  if (strings.size) {
    const items = [...strings.keys()].map((s) => `<si><t${needsPreserve(s) ? ' xml:space="preserve"' : ''}>${escapeText(s)}</t></si>`)
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
  const sheetList = book.sheets.map((s, i) => `<sheet name="${escapeAttr(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')
  // always there: without it Excel keeps the calculation settings it has from the books opened before (iteration on)
  const calc = `<calcPr${dynamic ? ' fullCalcOnLoad="1"' : ''}${iterationAttrs(book.iteration)}/>`
  const pr = book.dateSystem === 1904 ? '<workbookPr date1904="1"/>' : ''
  files.push({ name: 'xl/workbook.xml', text: `${HEAD}<workbook xmlns="${MAIN}" xmlns:r="${REL}">${pr}<bookViews><workbookView/></bookViews><sheets>${sheetList}</sheets>${definedNames(book, issues, tableNames)}${calc}</workbook>` })
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
    text: `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides}</Types>`,
  })
  const data = await zip(files.map((f) => ({ name: f.name, data: utf8(f.text) })))
  return { data, issues: issues.list() }
}
