import type { TextRun } from './contract.js'

/** A view's text runs are detached from the mutable workbook input. */
export const copyRuns = (runs: TextRun[]): TextRun[] => runs.map(r => ({ start: r.start, ...(r.font ? { font: { ...r.font } } : {}) }))
