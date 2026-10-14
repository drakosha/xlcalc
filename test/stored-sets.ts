/** Stored sets of Excel's answers read from `test/excel/`, without vitest (the bench's scripts read them too). */
import { existsSync, readdirSync, readFileSync } from 'node:fs'

/** The sets of a directory of `test/excel/` (`cond/`, `win/parts/`), by name, each with the folder of its saved files. */
export function storedSets<T>(sub: string): (T & { name: string; dir: URL })[] {
  const dir = new URL(`./excel/${sub}`, import.meta.url)
  if (!existsSync(dir)) return []
  return readdirSync(dir)
    .filter((f) => f.endsWith('.json'))
    .sort()
    .map((f) => ({ name: f.slice(0, -5), dir: new URL(`${f.slice(0, -5)}/`, dir), ...(JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as T) }))
}
