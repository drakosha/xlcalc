/**
 * `inspectKept`: what a host checks of a file without reading it into a workbook, in the file's own terms: the part of each
 * sheet, and the cells with formatted runs of text, which the workbook does not hold.
 */
import { object, onlyOptions } from '../../input/basic.js'
import { checkLimits, fileBytes } from '../read/book.js'
import { partText } from '../read/package.js'
import { attributes } from '../xml/scan.js'
import { CELL_XML, stringsWith, textWith } from './carriers.js'
import { openSource } from './source.js'
import type { InspectKeptOptions, KeptInspection } from './types.js'
import { childrenOf, rootOf } from './xml.js'

/** Formatted runs of text in the XML of a string; phonetic ones are `rPh`, which it does not take. */
const RUNS = /<(?:[\w.-]+:)?r[\s>]/

/** Reads what a host checks of a file without its workbook: its sheets' parts and the cells with formatted runs of text. */
export async function inspectKept(data: Uint8Array | ArrayBuffer, opts?: InspectKeptOptions): Promise<KeptInspection> {
  const bytes = fileBytes(data)
  const o = opts === undefined ? {} : object(opts, 'opts', 'options such as { limits }')
  onlyOptions(o, 'opts', ['limits'])
  const src = await openSource(bytes, o.limits === undefined ? undefined : checkLimits(o.limits))
  const strings = stringsWith(await partText(src.archive, src.wbRels, 'sharedStrings'), RUNS)
  const sheets = src.sheets.map((s) => {
    const cells = childrenOf(s.xml, rootOf(s.xml)!).find((c) => c.local === 'sheetData')
    const formattedText: string[] = []
    if (cells)
      for (const m of s.xml.slice(cells.open, cells.end).matchAll(CELL_XML)) {
        const a = attributes(m[1])
        if (a.r !== undefined && textWith(a, m[2] ?? '', strings, RUNS)) formattedText.push(a.r)
      }
    return { name: src.decls[s.decl].name, part: s.part, formattedText }
  })
  return { sheets }
}
