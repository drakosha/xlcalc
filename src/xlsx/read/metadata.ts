/**
 * Cell metadata (`xl/metadata.xml`): `cm` marks a dynamic array formula (XLDAPR metadata with `fDynamic`), `vm`
 * points at a rich value. Excel 365 stores #SPILL! and #CALC! that way: the cell holds `#VALUE!` for older readers
 * and a rich value of the `_error` structure gives the error's type.
 */
import { attributes, element, elements, flag } from '../xml/scan.js'

/** The errors of rich values by their `errorType`. */
const RICH_ERRORS = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CONNECT!', '#BLOCKED!', '#UNKNOWN!', '#FIELD!', '#CALC!']

export interface CellMetadata {
  /** Whether a cell metadata index (`cm`, 1-based) marks a dynamic array formula. */
  dynamic(cm: number): boolean
  /** The text of the error a value metadata index (`vm`, 1-based) stands for; undefined when it is no error (a picture, a data type). */
  error(vm: number): string | undefined
}

/** The metadata blocks of a kind: each block's first `rc` as type (1-based, into the metadata types) and value index. */
const blocks = (xml: string, local: string) => elements(element(xml, local)?.body ?? '', 'bk').map((bk) => attributes(element(bk.body ?? '', 'rc')?.attrs ?? ''))

/** The future metadata blocks of a type by name. */
const future = (xml: string, name: string) => elements(elements(xml, 'futureMetadata').find((e) => attributes(e.attrs).name === name)?.body ?? '', 'bk')

/** The rich values' errors by index: `rdrichvalue.xml` values in the order of their structure's keys. */
function richValueErrors(values: string | undefined, structures: string | undefined): (string | undefined)[] {
  if (!values || !structures) return []
  const keys = elements(structures, 's').map((s) => ({ error: attributes(s.attrs).t === '_error', keys: elements(s.body ?? '', 'k').map((k) => attributes(k.attrs).n) }))
  return elements(values, 'rv').map((rv) => {
    const s = keys[Number(attributes(rv.attrs).s)]
    if (!s?.error) return undefined
    const at = s.keys.indexOf('errorType')
    const type = Number(elements(rv.body ?? '', 'v')[at]?.body)
    return RICH_ERRORS[type] ?? '#UNKNOWN!'
  })
}

export function cellMetadata(xml: string | undefined, richValues?: string, richStructures?: string): CellMetadata {
  if (!xml) return { dynamic: () => false, error: () => undefined }
  const types = elements(element(xml, 'metadataTypes')?.body ?? '', 'metadataType').map((e) => attributes(e.attrs).name)
  const dynamicFlags = future(xml, 'XLDAPR').map((bk) => {
    const p = element(bk.body ?? '', 'dynamicArrayProperties')
    return p !== undefined && flag(attributes(p.attrs).fDynamic)
  })
  const richIndex = future(xml, 'XLRICHVALUE').map((bk) => Number(attributes(element(bk.body ?? '', 'rvb')?.attrs ?? '').i))
  const errors = richValueErrors(richValues, richStructures)
  const cells = blocks(xml, 'cellMetadata')
  const values = blocks(xml, 'valueMetadata')
  return {
    dynamic(cm) {
      const rc = cells[cm - 1]
      return rc !== undefined && types[Number(rc.t) - 1] === 'XLDAPR' && dynamicFlags[Number(rc.v)] === true
    },
    error(vm) {
      const rc = values[vm - 1]
      return rc !== undefined && types[Number(rc.t) - 1] === 'XLRICHVALUE' ? errors[richIndex[Number(rc.v)]] : undefined
    },
  }
}
