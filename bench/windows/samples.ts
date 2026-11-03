/**
 * The samples of the Number Format list on the Home tab of Excel for Windows over the cases of a samples set
 * (`bench/coverage/samples/<set>.ts`): the list exists only on a screen, so `samples.ps1` runs as a task in the session
 * of the machine's user (`screen.ts`), selects each case's cell, opens the list and reads its items through MSAA, the
 * sample of each item being its description. `npx tsx bench/windows/run.ts samples <set> [locale]` stores, for each
 * case, what the box showed, the cell's text and code (en-US), and the list's items with their samples in
 * `test/excel/win/samples/[<tag>/]<set>.json`; screenshots of the lists go to `bench/windows/out/samples-<set>-<tag>/`.
 */
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { writeXlsx } from '../../src/xlsx.js'
import { localeDir } from '../coverage/locales.js'
import { sampleAddress, sampleBook, sampleBooks, type SampleCase, type SampleSet } from '../coverage/samples/cases.js'
import { secondsSince, setModule, writeAnswers } from '../coverage/stored.js'
import { winFile } from './functions.js'
import { HOST, metaOf, OUT, type Tag, type WinEnv } from './remote.js'
import { runOnScreen, scp } from './screen.js'

/** A case's answer: the box's category, the cell's text and code (en-US), the list's items as [name, sample]. */
export interface SampleAnswer {
  box: string
  text: string
  code: string
  items: [string, string][]
  error?: string
}

export interface StoredSamples {
  meta: Record<string, string | number>
  cases: SampleCase[]
  answers: Record<string, SampleAnswer>
}

export const samplesFile = (set: string, tag: Tag) => winFile(`samples/${localeDir(tag)}${set}.json`)

export async function stepSamples(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const t0 = Date.now()
  const cases = (await setModule<SampleSet>('samples', set)).buildSampleCases()
  const books = sampleBooks(cases)
  const dir = mkdtempSync(join(tmpdir(), 'xlwin-samples-'))
  const job = []
  for (const [i, part] of books.entries()) {
    const { data, issues } = await writeXlsx({ book: sampleBook(part) }, { values: 'omit' })
    if (issues.length) throw new Error(`samples ${set}: the book does not go into a file whole: ${JSON.stringify(issues)}`)
    writeFileSync(join(dir, `b${i}.xlsx`), data)
    job.push({ file: `b${i}.xlsx`, cells: part.map((_, k) => sampleAddress(k)) })
  }
  writeFileSync(join(dir, 'job.json'), JSON.stringify(job))
  await runOnScreen('samples', 'samples.ps1', dir, 60 + 8 * cases.length)
  scp(`${HOST}:C:/xlcalc/samples/out.json`, join(dir, 'out.json'))
  const shots = join(OUT, `samples-${set}-${tag}`)
  rmSync(shots, { recursive: true, force: true })
  mkdirSync(OUT, { recursive: true })
  scp('-r', `${HOST}:C:/xlcalc/samples/shots`, shots)
  const raw = JSON.parse(readFileSync(join(dir, 'out.json'), 'utf8').replace(/^\uFEFF/, '')) as { version: string; ui: number; screen: string; box: string; error?: string; answers: (SampleAnswer & { cell: string; how: string; size: string })[][] }
  if (raw.error) throw new Error(`samples: ${raw.error}`)
  if (raw.answers.length !== books.length) throw new Error(`samples: ${raw.answers.length} books answered of ${books.length}`)
  const answers: Record<string, SampleAnswer> = {}
  books.forEach((part, i) =>
    part.forEach((c, k) => {
      const { box, text, code, items, error } = raw.answers[i][k]
      answers[c.id] = { box, text, code, items, ...(error ? { error } : {}) }
    }),
  )
  const meta = { ...metaOf(e, tag), box: raw.box, screen: raw.screen, date: new Date().toISOString().slice(0, 10), seconds: secondsSince(t0) }
  writeAnswers(samplesFile(set, tag), { meta, cases, answers })
  console.log(`samples ${set} ${tag}: ${Object.keys(answers).length} cells, ${Object.values(answers).filter((a) => a.error).length} with an error; screenshots in ${shots}`)
}
