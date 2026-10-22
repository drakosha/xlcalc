/**
 * The engine's answers to a held-out workbook set as its predictions (values by `Sheet!A1`, shown text by
 * `show:Sheet!A1`), written before Excel runs the set:
 * `npx tsx bench/coverage/workbook/predict.ts <set>` writes `workbook/<set>.predicted.json`, which the set reads.
 */
import { fileURLToPath } from 'node:url'
import { setModule, writeAnswers } from '../stored.js'
import { cellsOf } from '../case-book.js'
import type { WorkbookSet } from './cases.js'
import { engineReads } from '../../../test/workbook-cases.js'
import { shownValue } from '../../../test/spill-shown.js'

const set = process.argv[2]
const { buildWorkbookCases } = await setModule<WorkbookSet>('workbook', set)
const out: Record<string, Record<string, unknown>[]> = {}
for (const c of buildWorkbookCases()) {
  let reads
  try {
    reads = engineReads(c)
  } catch (e) {
    // a case the engine does not compute is predicted as its error
    out[c.id] = [{ thrown: String(e) }]
    continue
  }
  out[c.id] = reads.map((r) => {
    const vals: Record<string, unknown> = {}
    for (const g of c.read) for (const a of cellsOf(g.range).addrs) {
      const v = shownValue(r, g.sheet, a)
      if (v) vals[`${g.sheet}!${a}`] = v
    }
    // the text of the cells to show, by `show:Sheet!A1`
    for (const g of c.show ?? []) for (const a of cellsOf(g.range).addrs) vals[`show:${g.sheet}!${a}`] = r.sheets.find((s) => s.name === g.sheet)!.cells[a]?.text ?? ''
    return vals
  })
}
writeAnswers(fileURLToPath(new URL(`./${set}.predicted.json`, import.meta.url)), out)
