/**
 * The Sort command against Excel for Windows en-US: the book of `bench/coverage/sort/<set>.ts`, each case's column
 * sorted by itself without a header (`Range.Sort`, the step `sort` of `xl.ps1`), and the column's formulas read back
 * into `test/excel/win/sort/<set>.json`, as the Mac's step (`bench/coverage/sort.ts`) stores them.
 */
import { writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { writeXlsx } from '../../src/xlsx.js'
import { ROOT } from '../coverage/paths.js'
import { secondsSince, setModule, writeAnswers } from '../coverage/stored.js'
import { rangeOf, sortBook, type SortSet } from '../coverage/sort/cases.js'
import { inJobs, metaOf, type Tag, type WinEnv } from './remote.js'

export async function stepSort(name: string, tag: Tag, e: WinEnv): Promise<void> {
  const { set } = await setModule<{ set: SortSet }>('sort', name)
  const answers: Record<string, string[]> = {}
  const t0 = Date.now()
  const n = set.cases.length
  await inJobs(
    'sort',
    [set],
    async (s, file) => {
      writeFileSync(file, (await writeXlsx({ book: sortBook(s) }, { values: 'omit' })).data)
      const range = (i: number) => ({ sheet: 'S', range: rangeOf(s, i) })
      return { file, steps: [{ do: 'full' }, ...s.cases.map((c, i) => ({ do: 'sort', ...range(i), descending: !!c.descending })), ...s.cases.map((_, i) => ({ do: 'formulas', ...range(i) }))] }
    },
    (s, a) => {
      if (!a || !('answers' in a)) throw new Error(`Excel did not sort the book: ${a && 'rejected' in a ? a.rejected : 'no answer'}`)
      s.cases.forEach((c, i) => (answers[c.id] = (a.answers[1 + n + i] as string[][]).map(([f]) => f)))
      return []
    },
    () => {},
  )
  writeAnswers(join(ROOT, 'test/excel/win/sort', `${name}.json`), { meta: { ...metaOf(e, tag), seconds: secondsSince(t0) }, answers })
  console.log(`sort ${name}: ${n} cases`)
}
