/** Values made once by key, in a map that starts over when it reaches `limit` keys (a column repeats a few keys). */
export function remembered<K, T>(cache: Map<K, T>, key: K, limit: number, make: () => T): T {
  const hit = cache.get(key)
  if (hit !== undefined) return hit
  const made = make()
  if (cache.size >= limit) cache.clear()
  cache.set(key, made)
  return made
}
