/** Range.Copy and Get-Clipboard -Raw over SSH; every case runs without a screen, in its own book. */
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { writeXlsx } from '../../src/xlsx.js'
import { clipboardCases } from '../coverage/clipboard/cases.js'
import { localeDir } from '../coverage/locales.js'
import { writeAnswers } from '../coverage/stored.js'
import { metaOf, OUT, runJob, type Tag, type WinEnv } from './remote.js'

export async function stepClipboard(tag: Tag, env: WinEnv): Promise<void> {
  mkdirSync(OUT, { recursive: true })
  const cases = clipboardCases(tag)
  const books = await Promise.all(cases.map(async (c) => {
    const file = join(OUT, `clipboard-${c.id}.xlsx`)
    const { data, issues } = await writeXlsx({ book: c.book }, { values: 'omit' })
    if (issues.length) throw new Error(JSON.stringify(issues))
    writeFileSync(file, data)
    return { file, steps: [{ do: 'full' }, ...(c.steps ?? []), { do: 'copyText', sheet: 'S', range: c.range }, { do: 'shown', sheet: 'S', cells: ['A1', 'A2', 'A3'] }] }
  }))
  const got = await runJob(`clipboard-${tag}`, books)
  const answers = Object.fromEntries(cases.map((c, i) => {
    const a = got[i]
    if (!a || !('answers' in a)) throw new Error(`${c.id}: ${JSON.stringify(a)}`)
    return [c.id, { text: a.answers[1 + (c.steps?.length ?? 0)], shown: a.answers.at(-1) }]
  }))
  const file = join('test/excel/win/clipboard', localeDir(tag), 'copy-held.json')
  mkdirSync(join('test/excel/win/clipboard', localeDir(tag)), { recursive: true })
  writeAnswers(file, { meta: metaOf(env, tag), cases, answers })
  console.log(`${tag}: ${cases.length} clipboard cases recorded`)
}
