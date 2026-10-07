/**
 * `writeXlsx`: a book as an xlsx file, with formulas and their values or with values only. The package holds the
 * workbook, its sheets, styles, shared strings, the dynamic array metadata when there are formulas, and the
 * book's locale as a custom document property (xlsx has no place for it). A formula file asks Excel to compute
 * the book on opening, so Excel shows its own values over the ones written.
 */
import type { BookInput, ComputeResult } from '../../contract.js'
import { checkBook, InputError } from '../../input.js'
import { FormulaError } from '../../syntax/ast.js'
import { parse } from '../../syntax/parse.js'
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

const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
const TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml'

/** The dynamic array metadata (XLDAPR) that `cm="1"` points at, as Excel 365 writes it. */
const METADATA = `${HEAD}<metadata xmlns="${MAIN}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>`

const rels = (items: [type: string, target: string][]) =>
  `${HEAD}<Relationships xmlns="${PKG_REL}">${items.map(([type, target], i) => `<Relationship Id="rId${i + 1}" Type="${type}" Target="${target}"/>`).join('')}</Relationships>`

/** A name the file holds: Excel's rules (letters of any script, `_` or `\` first; not a cell address, R1C1 or a boolean). */
const NAME = /^[\p{L}_\\][\p{L}\p{N}_.\\?]*$/u
const NOT_NAME = /^(?:[A-Z]{1,3}\d+|R\d*C\d*|R|C|TRUE|FALSE)$/i

function definedNames(book: BookInput, issues: Issues): string {
  const index = new Map(book.sheets.map((s, i) => [sheetKey(s.name), i]))
  const seen = new Set<string>()
  const out: string[] = []
  for (const n of book.names ?? []) {
    const scope = n.scope === undefined ? undefined : index.get(sheetKey(n.scope))
    const key = `${scope ?? ''}\n${n.name.toUpperCase()}`
    let text: string | undefined
    try {
      text = toFile(parse(n.ref))
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
              ? 'a name whose reference does not parse'
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
  const sheets = book.sheets.map((s, i) => writeSheet(s, i, { styles, string, issues, sheets: sheetKeys, mode, result: cells?.[i] }))
  const dynamic = sheets.some((s) => s.formulas)
  const files: { name: string; text: string }[] = []
  const contentTypes: [string, string][] = [['/xl/workbook.xml', `${TYPE}.sheet.main+xml`]]
  const wbRels: [string, string][] = []
  sheets.forEach((s, i) => {
    files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, text: s.xml })
    contentTypes.push([`/xl/worksheets/sheet${i + 1}.xml`, `${TYPE}.worksheet+xml`])
    wbRels.push([`${REL}/worksheet`, `worksheets/sheet${i + 1}.xml`])
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
  if (dynamic) {
    files.push({ name: 'xl/metadata.xml', text: METADATA })
    contentTypes.push(['/xl/metadata.xml', `${TYPE}.sheetMetadata+xml`])
    wbRels.push([`${REL}/sheetMetadata`, 'metadata.xml'])
  }
  const sheetList = book.sheets.map((s, i) => `<sheet name="${escapeAttr(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')
  const calc = dynamic ? '<calcPr fullCalcOnLoad="1"/>' : ''
  files.push({ name: 'xl/workbook.xml', text: `${HEAD}<workbook xmlns="${MAIN}" xmlns:r="${REL}"><bookViews><workbookView/></bookViews><sheets>${sheetList}</sheets>${definedNames(book, issues)}${calc}</workbook>` })
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
