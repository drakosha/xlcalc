/** The xlsx filter files (`bench/coverage/xlsx/filters.ts`) against a run of Excel that computed and saved them. */
import { expect, it } from 'vitest'
import type { StoredRun } from '../bench/coverage/xlsx.js'
import { FILTER_FILES, filterFile, type FilterFile } from '../bench/coverage/xlsx/filters.js'
import { partsHash } from '../bench/coverage/xlsx/sent.js'

/** Every file went through the run of `excel` and only those, each as sent; `check` checks each file as it came back. */
export function filterRun(run: StoredRun, excel: string, check: (c: FilterFile) => Promise<void>): void {
  it(`every file of the set went through ${excel}, and only those`, () => {
    expect(Object.keys(run.books).sort()).toEqual(FILTER_FILES.map((c) => c.id).sort())
  })
  for (const c of FILTER_FILES)
    it(`${c.id}: ${c.about}`, async () => {
      expect(run.books[c.id].sent, 'the file changed since Excel saw it: run the xlsx filters step').toBe(await partsHash(await filterFile(c)))
      await check(c)
    })
}
