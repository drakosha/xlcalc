/** A result's diagnostics of one kind, without the field `kind`: the cells with an error, the names with a problem. */
import type { CellIssue, ComputeResult, NameIssue } from '../src/index.js'

export const cellIssues = (r: Pick<ComputeResult, 'diagnostics'>): CellIssue[] => r.diagnostics.flatMap(({ kind, ...d }) => (kind === 'cell' ? [d as CellIssue] : []))

export const nameIssues = (r: Pick<ComputeResult, 'diagnostics'>): NameIssue[] => r.diagnostics.flatMap(({ kind, ...d }) => (kind === 'name' ? [d as NameIssue] : []))
