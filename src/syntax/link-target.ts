/**
 * A link's target and the file's name a person reads: a file stores the target as an address (`Report%202025.xlsx`), a
 * formula shows and types the name (`[Report 2025.xlsx]`). Measured in Excel for Windows (`bench/coverage/links/names.ts`).
 */
import { checkString } from '../input/basic.js'

/** A path without `file://`, split into its folder (`C:\\data\\`, `../data/`, empty for a file name) and its file. */
export function splitPath(path: string): { dir: string; file: string } {
  const t = path.replace(/^file:\/\/\/?/i, '')
  const k = Math.max(t.lastIndexOf('/'), t.lastIndexOf('\\')) + 1
  return { dir: t.slice(0, k), file: t.slice(k) }
}

/** A target with its escapes undone (`%20` a space, `%D0%9A` `К`), as Excel shows it; a `%` that escapes nothing stays. */
function unescaped(target: string): string {
  return target.replace(/(?:%[0-9a-f]{2})+/gi, (run) => {
    try {
      return decodeURIComponent(run)
    } catch {
      return run
    }
  })
}

/** A link's target as a formula shows it: the folder and the file, unescaped. */
export const shownTarget = (target: string) => splitPath(unescaped(target))

/**
 * The file's name of a link's target as a formula shows it, by which the package finds a link: without its folder and its
 * escapes (`file:///C:\\data\\Report%202025.xlsm` is `Report 2025.xlsm`). Compare two with `sameName`.
 */
export function linkFileName(target: string): string {
  return shownTarget(checkString(target, 'target')).file
}

/**
 * The target of a link to a file by its name or path, as Excel writes it: `%`, a space, `#` and `?` escaped, the rest as it
 * is (Cyrillic, `&`, a path's separators), so `linkFileName` gives the name back. `?`, which Windows does not allow in a
 * file's name, follows `#` and was not measured.
 */
export function linkTarget(name: string): string {
  return checkString(name, 'name').replace(/[% #?]/g, (c) => `%${c.charCodeAt(0).toString(16).toUpperCase().padStart(2, '0')}`)
}
