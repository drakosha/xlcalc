/**
 * The xlsx tests of IronCalc in Excel for Windows: each file of `bench/coverage/ironcalc/cases.json`, from a checkout of
 * IronCalc (`IRONCALC_ROOT`, at the cases' commit), opened as it is, computed fully (`CalculateFull`), and the `Value2` of
 * each case's cell read. The files stay IronCalc's: none is written here or stored. The answers go to
 * `bench/coverage/ironcalc/win.json`, beside the value each file stored and xlcalc's.
 */
import { copyFileSync, existsSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { secondsSince } from '../coverage/stored.js'
import { valOf, type WinCell } from './functions.js'
import { again, inJobs, metaOf, type Tag, type WinEnv } from './remote.js'

const HERE = join(fileURLToPath(new URL('.', import.meta.url)), '../coverage/ironcalc')

interface Case {
  id: string
  file: string
  cell: string
}

/** Excel's value as the cases keep values: `{ t, v }`, `t` one of number, string, boolean, e, empty. */
function asCase(c: WinCell): { t: string; v?: unknown } {
  const v = valOf(c)
  if (!v) return { t: 'empty' }
  return { t: ({ n: 'number', s: 'string', b: 'boolean' } as Record<string, string>)[v.t] ?? v.t, v: v.v }
}

export async function stepIronCalc(tag: Tag, e: WinEnv): Promise<void> {
  const root = process.env.IRONCALC_ROOT
  if (!root || !existsSync(join(root, 'xlsx/tests'))) throw new Error('IRONCALC_ROOT: a checkout of IronCalc at the commit of bench/coverage/ironcalc/cases.json')
  const { cases } = JSON.parse(readFileSync(join(HERE, 'cases.json'), 'utf8')) as { cases: Case[] }
  const out = join(HERE, 'win.json')
  const answers: Record<string, unknown> = existsSync(out) ? (JSON.parse(readFileSync(out, 'utf8')) as { answers: Record<string, unknown> }).answers : {}
  const rejected: string[] = []
  const t0 = Date.now()
  const files = [...new Set(cases.filter((c) => !answers[c.id]).map((c) => c.file))]
  const ofFile = (f: string) => cases.filter((c) => c.file === f)
  const save = () => writeFileSync(out, JSON.stringify({ meta: { ...metaOf(e, tag), seconds: secondsSince(t0) }, answers, ...(rejected.length ? { rejected } : {}) }, null, 1) + '\n')
  await inJobs(
    'ironcalc',
    files.map((f) => [f]),
    async ([f], file) => {
      copyFileSync(join(root, 'xlsx/tests', f), file)
      return { file, steps: [{ do: 'full' }, ...ofFile(f).map((c) => { const [sheet, range] = c.cell.split('!'); return { do: 'values', sheet, range } })] }
    },
    (part, a, cutShort) => {
      if (a && 'answers' in a) {
        ofFile(part[0]).forEach((c, i) => (answers[c.id] = asCase((a.answers[i + 1] as WinCell[][])[0][0])))
        return []
      }
      return again(part, a, cutShort, (f) => rejected.push(f))
    },
    save,
    10,
  )
  save()
}
