/**
 * Stored answers in a stable order: a file of answers lists them in the order of its cases, so a run that adds or
 * reruns cases changes the lines of those cases only, however the run went.
 */

/** The answers of `rec` in the order of `cases` (ids without an answer left out). */
export const inCaseOrder = <T>(cases: readonly { id: string }[], rec: Record<string, T>): Record<string, T> =>
  Object.fromEntries(cases.flatMap((c) => (c.id in rec ? [[c.id, rec[c.id]]] : [])))
