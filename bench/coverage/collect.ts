/** Cases collected by group, as the case sets build them. */

/** Cases of groups, ids `<group>-<n>`: `add(group, c)` puts a case with the next number of its group. */
export function groupCollector<C extends { id: string }>(): { out: C[]; add: (group: string, c: Omit<C, 'id'>) => void } {
  const out: C[] = []
  const count = new Map<string, number>()
  const add = (group: string, c: Omit<C, 'id'>) => {
    const n = (count.get(group) ?? 0) + 1
    count.set(group, n)
    out.push({ id: `${group}-${n}`, ...c } as C)
  }
  return { out, add }
}
