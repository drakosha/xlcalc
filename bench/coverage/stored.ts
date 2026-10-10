/**
 * Stored answers in a stable order: a file of answers lists them in the order of its cases, so a run that adds or
 * reruns cases changes the lines of those cases only, however the run went.
 */

/** The answers of `rec` in the order of `cases` (ids without an answer left out). */
export const inCaseOrder = <T>(cases: readonly { id: string }[], rec: Record<string, T>): Record<string, T> =>
  Object.fromEntries(cases.flatMap((c) => (c.id in rec ? [[c.id, rec[c.id]]] : [])))

/** A stored file of answers to cases sent as files: what was sent for each case (`partsHash`), the answers, the books Excel did not load. */
export interface StoredAnswers<C extends { id: string }, A> {
  cases: C[]
  sent: Record<string, string>
  answers: Record<string, A>
  rejected?: Record<string, string>
}

/** The stored answers a run keeps: a case's answer (its saved file there too) or refusal, while the case and its file are the same. */
export function keptAnswers<C extends { id: string }, A>(cases: C[], sent: Record<string, string>, prev: StoredAnswers<C, A> | null, saved: (id: string) => boolean): { answers: Record<string, A>; rejected: Record<string, string> } {
  const key = (c: C, hash: string | undefined) => JSON.stringify([c, hash])
  const same = new Map(prev?.cases.map((c) => [c.id, key(c, prev.sent[c.id])]) ?? [])
  const answers: Record<string, A> = {}
  const rejected: Record<string, string> = {}
  for (const c of cases)
    if (same.get(c.id) === key(c, sent[c.id])) {
      if (prev!.answers[c.id] && saved(c.id)) answers[c.id] = prev!.answers[c.id]
      else if (prev!.rejected?.[c.id]) rejected[c.id] = prev!.rejected[c.id]
    }
  return { answers, rejected }
}
