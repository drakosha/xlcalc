/** Stored sets of Excel's answers read from `test/excel/`, without vitest (the bench's scripts read them too). */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { localeDir, TAGS } from '../bench/coverage/locales.js'

/** The sets of a directory of `test/excel/` (`cond/`, `win/parts/`), by name, each with the folder of its saved files. */
export function storedSets<T>(sub: string, keep: (f: string) => boolean = (f) => f.endsWith('.json')): (T & { name: string; dir: URL })[] {
  const dir = new URL(`./excel/${sub}${sub.endsWith('/') ? '' : '/'}`, import.meta.url)
  if (!existsSync(dir)) return []
  return readdirSync(dir)
    .filter(keep)
    .sort()
    .map((f) => ({ name: f.slice(0, -5), dir: new URL(`${f.slice(0, -5)}/`, dir), ...(JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as T) }))
}

/** The sets of a kind of `test/excel/` in every locale Excel ran it in (`entry/`, `win/end/`), each with its tag. */
export function localeSets<T>(sub: string, keep: (f: string) => boolean = (f) => f.endsWith('.json')): { tag: string; set: string; stored: T }[] {
  return TAGS.flatMap((tag) => {
    const dir = new URL(`./excel/${sub}/${localeDir(tag)}`, import.meta.url)
    if (!existsSync(dir)) return []
    return readdirSync(dir).filter(keep).sort().map((f) => ({ tag, set: f.slice(0, -5), stored: JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as T }))
  })
}
