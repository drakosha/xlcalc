/** The names a session's edit takes (`setNames`, an operation `setNames`): the workbook's names list, as it replaces it. */
import type { NameInput } from '../contract.js'
import type { EngineState } from '../engine.js'
import { InputError } from '../input-error.js'
import { localeBooleans } from '../locale/registry.js'
import { prepareNames } from '../names.js'
import { tableNames } from '../tables/model.js'

/**
 * Refuses, with `InputError` at `path[i].name`, a name the workbook would drop (`NameIssue.dropped`, as `validateNames`
 * finds it), unless the workbook has it already as it stands: the list comes back whole, with the names it was opened
 * with (a hidden one among them) unchanged.
 */
export function checkNamesKept(s: EngineState, names: NameInput[], path: string): void {
  const had = new Set((s.input.names ?? []).map(nameKey))
  const { issues } = prepareNames(s.sheetIdx, names, s.externalFunctions, tableNames(s.tables), localeBooleans(s.input.locale), s.input.externalBooks)
  for (const issue of issues) {
    if (!issue.dropped) continue
    // the last of its name and scope: a duplicate drops the later one
    const i = names.findLastIndex((n) => n.name === issue.name && n.scope === issue.scope)
    if (!had.has(nameKey(names[i]))) throw new InputError(`${path}[${i}].name`, issue.reason)
  }
}

const nameKey = (n: NameInput): string => JSON.stringify([n.name, n.scope, n.ref, n.hidden])
