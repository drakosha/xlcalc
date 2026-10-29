/**
 * The stored answers of the entry sets in each locale of `LOCALES` (`test/excel/entry/`, `test/excel/win/entry/`, from
 * `run.ts entry <set> [locale]` on the Mac and `bench/windows/run.ts entry <set> [locale]`), and the check of one case:
 * the stored case is its generator's current one, and `parseEntry` in the locale, on the day of the run (a date typed
 * without a year takes its year), into a cell of the case's format, gives Excel's answer, or differs where it is known to;
 * the number Excel's object model gives as the cell's value, where the run read it, is the one its file holds.
 */
import { expect } from 'vitest'
import type { StoredEntry } from '../bench/coverage/entry.js'
import type { EntryCase, EntrySet } from '../bench/coverage/entry/cases.js'
import { engineEntry, entryDifferences } from '../bench/coverage/entry/compare.js'
import { ENTRY_REASONS, type EntryReason } from './entry-known.js'
import { localeSets } from './stored-sets.js'
import './locales.js'

export interface EntrySetFile extends StoredEntry {
  name: string
  tag: string
}

/** The sets stored under `test/excel/<root>entry/`, every locale's. */
export function entrySets(root: '' | 'win/'): EntrySetFile[] {
  return localeSets<StoredEntry>(`${root}entry`).map((r) => ({ name: r.set, tag: r.tag, ...r.stored }))
}

const generators = new Map<string, Promise<EntrySet>>()

/** A set's generator. */
const generator = (name: string) => {
  let g = generators.get(name)
  if (!g) generators.set(name, (g = import(`../bench/coverage/entry/${name}.ts`) as Promise<EntrySet>))
  return g
}

/** One stored case against `parseEntry`: Excel's answer, or a difference where `known` names a reason. */
export async function checkCase(set: EntrySetFile, c: EntryCase, known: (key: string) => EntryReason | undefined): Promise<void> {
  const mod = await generator(set.name)
  expect(mod.buildEntryCases().find((x) => x.id === c.id), "the stored case is the generator's current one").toEqual(c)
  const got = engineEntry(c, set.tag, String(set.meta.date), mod.data)
  const want = set.answers[c.id]
  if (want.value2 !== undefined && typeof want.cell?.value === 'number') expect(want.value2, "the object model's value is the number the file holds").toBe(want.cell.value)
  const diff = entryDifferences(got, want, set.tag, c.text)
  const reason = known(`${set.tag}:${set.name}:${c.id}`)
  if (reason) expect(diff.length, `${c.id} is known to differ: ${ENTRY_REASONS[reason]}`).toBeGreaterThan(0)
  else expect(diff).toEqual([])
}
