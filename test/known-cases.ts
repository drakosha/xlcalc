/** The shared shape of the tests' known-difference tables: `<tag>:<set>` and its case ids → the same reason. */
export function knownCases<R extends string>(reason: R, bySet: Record<string, string[]>): [string, R][] {
  return Object.entries(bySet).flatMap(([set, ids]) => ids.map((id): [string, R] => [`${set}:${id}`, reason]))
}
