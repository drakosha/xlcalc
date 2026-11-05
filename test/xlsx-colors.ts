/** The fill colours of a hand-written book against what an Excel showed, for `xlsx-excel.test.ts` and `xlsx-win.test.ts`. */
import { readFileSync } from 'node:fs'
import { colorBook, type ColorCase } from '../bench/coverage/xlsx/colors.js'
import { readXlsx } from '../src/xlsx.js'
import { cellsOf } from './looks-book.js'

/** The cases where the reader's colour is not Excel's. */
export async function filledAs(cases: ColorCase[], stored: Record<string, string>, theme?: string): Promise<string[]> {
  const r = await readXlsx(await colorBook(cases, theme))
  return cases.flatMap((c) => {
    const cell = cellsOf(r.book.sheets[0])[c.cell]
    const got = cell?.style === undefined ? 'none' : r.book.styles![cell.style].fill!.color
    return got === stored[c.color] ? [] : [`${c.color}: Excel ${stored[c.color]}, reader ${got}`]
  })
}

/** The cases where the prediction committed before the run (`bench/coverage/xlsx/colors-held.predicted.json`) is not Excel's. */
export function heldOff(stored: Record<string, string>): string[] {
  const { colors } = JSON.parse(readFileSync(new URL('../bench/coverage/xlsx/colors-held.predicted.json', import.meta.url), 'utf8')) as { colors: Record<string, string> }
  return Object.entries(colors).flatMap(([k, v]) => (stored[k] === v ? [] : [`${k}: Excel ${stored[k]}, predicted ${v}`]))
}
