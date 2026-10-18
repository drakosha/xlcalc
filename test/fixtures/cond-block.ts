import type { BookInput } from '../../src/contract.js'

/** Sheet S: a block of `rows` rows by 10 values (A to J) with one expression rule `formula` over all of it, bold where it holds. */
export function condBlock(rows: number, formula: string): BookInput {
  const cells: Record<string, { value: number }> = {}
  for (let r = 1; r <= rows; r++) for (const c of 'ABCDEFGHIJ') cells[`${c}${r}`] = { value: (r * 7 + c.charCodeAt(0)) % 100 }
  return { sheets: [{ name: 'S', cells, conditionalFormats: [{ type: 'expression', ranges: [`A1:J${rows}`], formula, style: { font: { bold: true } } }] }] }
}

/** The windows of 50 rows of the block, from the top to the bottom. */
export const windows = (rows: number) => Array.from({ length: Math.ceil(rows / 50) }, (_, i) => `A${i * 50 + 1}:J${i * 50 + 50}`)
