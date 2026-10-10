/**
 * The stored answers of the entry sets in each locale of `LOCALES` (`test/excel/entry/`, `test/excel/win/entry/`, from
 * `run.ts entry <set> [locale]` on the Mac and `bench/windows/run.ts entry <set> [locale]`), and the check of one case:
 * the stored case is its generator's current one, and `parseEntry` in the locale, on the day of the run (a date typed
 * without a year takes its year), into a cell of the case's format, gives Excel's answer, or differs where it is known to.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { expect } from 'vitest'
import type { StoredEntry } from '../bench/coverage/entry.js'
import type { EntryCase, EntrySet } from '../bench/coverage/entry/cases.js'
import { entryDifferences } from '../bench/coverage/entry/compare.js'
import { localeDir, TAGS } from '../bench/coverage/locales.js'
import { parseEntry } from '../src/index.js'
import { ENTRY_REASONS, type EntryReason } from './entry-known.js'

export interface EntrySetFile extends StoredEntry {
  name: string
  tag: string
}

/** The sets stored under `test/excel/<root>entry/`, every locale's. */
export function entrySets(root: '' | 'win/'): EntrySetFile[] {
  return TAGS.flatMap((tag) => {
    const dir = new URL(`./excel/${root}entry/${localeDir(tag)}`, import.meta.url)
    if (!existsSync(dir)) return []
    return readdirSync(dir)
      .filter((f) => f.endsWith('.json'))
      .map((f) => ({ name: f.slice(0, -5), tag, ...(JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as StoredEntry) }))
  })
}

const generators = new Map<string, Promise<EntryCase[]>>()

/** A set's current cases, by its generator. */
const current = (name: string) => {
  let g = generators.get(name)
  if (!g) generators.set(name, (g = (import(`../bench/coverage/entry/${name}.ts`) as Promise<EntrySet>).then((m) => m.buildEntryCases())))
  return g
}

/** One stored case against `parseEntry`: Excel's answer, or a difference where `known` names a reason. */
export async function checkCase(set: EntrySetFile, c: EntryCase, known: (key: string) => EntryReason | undefined): Promise<void> {
  expect((await current(set.name)).find((x) => x.id === c.id), "the stored case is the generator's current one").toEqual(c)
  const got = parseEntry(c.text, { locale: set.tag, now: String(set.meta.date), ...(c.format !== undefined ? { cell: { format: c.format } } : {}) })
  const diff = entryDifferences(got, set.answers[c.id], set.tag)
  const reason = known(`${set.tag}:${set.name}:${c.id}`)
  if (reason) expect(diff.length, `${c.id} is known to differ: ${ENTRY_REASONS[reason]}`).toBeGreaterThan(0)
  else expect(diff).toEqual([])
}
