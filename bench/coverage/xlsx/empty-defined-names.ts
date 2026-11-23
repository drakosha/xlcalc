import type { BookInput } from '../../../src/contract.js'
import { compute } from '../../../src/index.js'
import { readXlsx, writeXlsx } from '../../../src/xlsx.js'
import { editParts } from './sent.js'

export interface EmptyNameCase {
  id: string
  /** Undefined leaves a valid name in place. Null uses a self-closing element. */
  body?: string | null
  local?: true
}

export const EMPTY_NAME_CASES: readonly EmptyNameCase[] = [
  { id: 'constant' },
  { id: 'broken-reference', body: '#REF!' },
  { id: 'empty', body: '' },
  { id: 'self-closing', body: null },
  { id: 'whitespace', body: ' \n\t ' },
  { id: 'local-shadow', body: '', local: true },
]

export async function emptyNameFile(c: EmptyNameCase): Promise<Uint8Array> {
  const names: BookInput['names'] = c.local
    ? [{ name: 'VacantName', ref: '17' }, { name: 'VacantName', scope: 'S', ref: '31', hidden: true }]
    : [{ name: 'VacantName', ref: '31' }]
  const book: BookInput = { names, sheets: [{ name: 'S', cells: { B1: { formula: 'VacantName' } } }] }
  const { data } = await writeXlsx({ book }, { values: 'omit' })
  if (c.body === undefined) return data
  return editParts(data, (part, xml) => {
    if (part !== 'xl/workbook.xml') return undefined
    const changed = xml.replace(/(<definedName\b[^>]*>)31<\/definedName>/, (_, tag: string) =>
      c.body === null ? tag.slice(0, -1) + '/>' : `${tag}${c.body}</definedName>`)
    if (changed === xml) throw new Error(`${c.id}: the workbook did not contain the expected defined name`)
    return changed
  })
}

export async function emptyNamePrediction(c: EmptyNameCase) {
  const { book, issues } = await readXlsx(await emptyNameFile(c))
  const result = compute(book)
  return { names: book.names, issues: issues.map(({ code, message }) => ({ code, message })), value: result.sheets[0].cells.B1.value }
}
