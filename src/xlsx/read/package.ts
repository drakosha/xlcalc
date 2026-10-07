/**
 * The package's relationships: a part's `.rels` lists the parts it points to. A target is relative to the
 * part's folder or absolute from the package root, may hold `..` and percent escapes, and is matched in any
 * case (the archive compares names case-insensitively). Kinds are the last segment of the relationship type,
 * the same in Transitional and Strict files (`…/relationships/worksheet`).
 */
import type { Archive } from '../zip/read.js'
import { attributes, elements } from '../xml/scan.js'
import { decodeXml } from '../xml/text.js'

export interface Rel {
  id: string
  /** The last segment of the type: `worksheet`, `styles`, `sharedStrings`, `officeDocument` … */
  kind: string
  /** The part's path in the package, or the target as written for an external one. */
  target: string
  external: boolean
}

/** `xl/worksheets` + `../theme/theme1.xml` → `xl/theme/theme1.xml`; absolute targets start from the root. */
function resolvePart(dir: string, target: string): string {
  let t = target
  try {
    t = decodeURIComponent(target)
  } catch {
    // a lone % is kept as written
  }
  const parts = t.startsWith('/') ? [] : dir.split('/').filter(Boolean)
  for (const seg of t.split('/')) {
    if (seg === '..') parts.pop()
    else if (seg && seg !== '.') parts.push(seg)
  }
  return parts.join('/')
}

const dirOf = (part: string) => part.slice(0, part.lastIndexOf('/') + 1)

/** The relationships of a part (`''` for the package itself); none when it has no `.rels`. */
export async function relsOf(archive: Archive, part: string): Promise<Rel[]> {
  const dir = dirOf(part)
  const xml = await archive.text(`${dir}_rels/${part.slice(dir.length)}.rels`)
  if (!xml) return []
  return elements(xml, 'Relationship').map((e) => {
    const a = attributes(e.attrs)
    const type = a.Type ?? ''
    const external = a.TargetMode === 'External'
    const target = decodeXml(a.Target ?? '')
    return { id: a.Id ?? '', kind: type.slice(type.lastIndexOf('/') + 1), target: external ? target : resolvePart(dir, target), external }
  })
}
