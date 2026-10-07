/**
 * Every stored file of Excel's answers lists them in the order of its cases (`stored.ts`), so a run that adds cases
 * changes only their lines.
 */
import { readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const ROOT = fileURLToPath(new URL('../../test/excel/', import.meta.url))
const files = (d: string): string[] => readdirSync(d, { withFileTypes: true }).flatMap((e) => (e.isDirectory() ? files(join(d, e.name)) : e.name.endsWith('.json') ? [join(d, e.name)] : []))
const measured = (JSON.parse(readFileSync(join(ROOT, 'coverage/cases.json'), 'utf8')) as { cases: { id: string }[] }).cases

describe('stored answers', () => {
  it('are in the order of their cases', () => {
    const off: string[] = []
    for (const f of files(ROOT)) {
      const j = JSON.parse(readFileSync(f, 'utf8')) as { cases?: { id: string }[]; values?: object; answers?: object }
      const rec = j.values ?? j.answers
      const cases = Array.isArray(j.cases) ? j.cases : /coverage[/\\]excel-/.test(f) ? measured : undefined
      if (!rec || !cases) continue
      const keys = Object.keys(rec)
      if (JSON.stringify(keys) !== JSON.stringify(cases.map((c) => c.id).filter((id) => id in rec))) off.push(f.slice(ROOT.length))
    }
    expect(off).toEqual([])
  })
})
