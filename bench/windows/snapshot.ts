/**
 * Excel for Windows's own formula text in a locale, beside the Mac's snapshot (`bench/locale/snapshot.ts`): the probe
 * books of `bench/locale/probes.ts`, each part's column A of the case sheet read as `Formula` (English) and
 * `FormulaLocal`, nothing calculated, into `test/excel/win/locale/<tag>.json` in the snapshot's shape. A part Excel
 * does not open is halved down to the formulas it refuses (`rejected`); a formula it deletes as it loads is `dropped`.
 * `international` is left empty: the answers' `meta` names the separators Windows computed with.
 */
import { mkdirSync, writeFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { writeBook } from '../coverage/book.js'
import type { Probe } from '../coverage/probes.js'
import { secondsSince } from '../coverage/stored.js'
import { serialize, type Snapshot } from '../locale/derive.js'
import { buildProbes, SHEETS } from '../locale/probes.js'
import { again, inJobs, metaOf, type Tag, type WinEnv } from './remote.js'

const stripEq = (f: unknown) => (typeof f === 'string' && f.startsWith('=') ? f.slice(1) : null)

export async function stepSnapshot(tag: Tag, e: WinEnv): Promise<void> {
  const { parts, functions } = buildProbes()
  const t0 = Date.now()
  const formulas: Snapshot['formulas'] = {}
  const rejected: string[] = []
  const dropped: string[] = []
  const file = fileURLToPath(new URL(`../../test/excel/win/locale/${tag}.json`, import.meta.url))
  const save = () => {
    const meta = { ...metaOf(e, tag), date: new Date().toISOString().slice(0, 10), seconds: secondsSince(t0) }
    mkdirSync(fileURLToPath(new URL('../../test/excel/win/locale/', import.meta.url)), { recursive: true })
    writeFileSync(file, serialize({ meta: meta as unknown as Snapshot['meta'], international: {}, functions, formulas, rejected, dropped }))
  }
  const book = async (part: Probe[], f: string) => {
    // a range of two cells at least reads as a grid
    const sent = part.length === 1 ? [...part, { id: 'pad', group: 'pad', formula: '0' }] : part
    await writeBook(sent, f, SHEETS)
    const range = `A1:A${sent.length}`
    return { file: f, steps: [{ do: 'formulas', sheet: 'C', range }, { do: 'formulasLocal', sheet: 'C', range }] }
  }
  const take = (part: Probe[], a: Parameters<typeof again>[1], cutShort: boolean): Probe[][] => {
    if (a && 'answers' in a) {
      const [en, loc] = a.answers as [unknown[][], unknown[][]]
      part.forEach((c, i) => {
        if (en[i][0] === '' && loc[i][0] === '') return void dropped.push(c.id)
        const pair = [stripEq(en[i][0]), stripEq(loc[i][0])]
        if (pair[0] === null || pair[1] === null) throw new Error(`${c.id} ${c.formula}: Excel gives ${JSON.stringify([en[i][0], loc[i][0]])}`)
        formulas[c.id] = pair as [string, string]
      })
      return []
    }
    return again(part, a, cutShort, (c) => rejected.push(c.id))
  }
  await inJobs(`snapshot-${tag}`, parts, book, take, save, 10)
  save()
  console.log(`${Object.keys(formulas).length} formulas, ${rejected.length} rejected, ${dropped.length} dropped → ${file}`)
}
