import { readXlsx, writeXlsx } from '../../../src/xlsx.js'
import { escapeAttr } from '../../../src/xlsx/xml/text.js'
import { editParts } from './sent.js'

export const FORMAT_CONDITION_CASES = [
  { id: 'three-conditions', code: '[>=1][Green]0.0%;[>=0.8][Blue]0.0%;[>=0.6][Magenta]0.0%;[Red]0.0%' },
  { id: 'two-conditions', code: '[>=1][Green]0.0%;[>=0.8][Blue]0.0%;[Magenta]0.0%;[Red]0.0%' },
  { id: 'three-sections', code: '[>=1][Green]0.0%;[>=0.8][Blue]0.0%;[>=0.6][Magenta]0.0%' },
  { id: 'only-third', code: '0.0%;[Blue]0.0%;[>=0.6][Magenta]0.0%' },
  { id: 'only-fourth', code: '0.0%;[Blue]0.0%;[Magenta]0.0%;[>=0.6]@' },
  { id: 'quoted-third', code: '0.0%;[Blue]0.0%;0.0%"[>=0.6]"' },
] as const

export async function conditionFormatFile(code: string): Promise<Uint8Array> {
  const cells = { A1: { value: 0.75, numberFormat: '0.00000' } }
  const { data } = await writeXlsx({ book: { sheets: [{ name: 'S', cells }] } })
  return editParts(data, (part, xml) => part === 'xl/styles.xml' ? xml.replace('formatCode="0.00000"', `formatCode="${escapeAttr(code)}"`) : undefined)
}

export async function conditionFormatPrediction(code: string) {
  const { book, issues } = await readXlsx(await conditionFormatFile(code))
  const imported = book.sheets[0].cells.A1.numberFormat
  const refusal = await writeXlsx({ book }).then(() => null, (error: Error) => error.message)
  return { imported, issues: issues.map(({ code, message }) => ({ code, message })), refusal }
}
