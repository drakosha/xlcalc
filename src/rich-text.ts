/** The text runs of a cell as a view gives them (`CellInput.richText`). */
import type { TextRun } from './contract.js'

/** Copies of runs, as a pool's are: a host that edits them does not change the workbook. */
export const copyRuns = (runs: TextRun[]): TextRun[] => runs.map((r) => ({ start: r.start, ...(r.font ? { font: { ...r.font } } : {}) }))
