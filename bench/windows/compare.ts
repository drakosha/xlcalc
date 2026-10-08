/**
 * Excel for Windows against Excel for Mac and the engine, case by case, over the formula answers stored in
 * `test/excel/win/` (`npx tsx bench/windows/compare.ts > out/compare.md`): a table of the cases where Windows and the
 * Mac part, with the engine's answer and which of the two it follows, then the cases where both agree and the engine
 * does not. Numbers compare as the tests compare them.
 */
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { Val } from '../coverage/formulajs.js'
import { matches, show } from '../../test/excel-cases.js'
import { runWinSet, WIN_SETS, type WinOutcome, type WinSet } from '../../test/excel-win-cases.js'

const ROOT = join(fileURLToPath(new URL('.', import.meta.url)), '../..')

/** The Mac's answers to the same cases, from the file the Mac harness stored. */
function macValues(set: WinSet): Record<string, Val> | null {
  const ru = set.tag === 'ru-RU'
  const rel = set.name === 'measurement' ? `coverage/excel-${ru ? 'ru' : 'en'}.json` : set.name === 'probes' ? 'probes.json' : `cases/${ru ? 'ru-RU/' : ''}${set.name}.json`
  const file = join(ROOT, 'test/excel', rel)
  return existsSync(file) ? (JSON.parse(readFileSync(file, 'utf8')) as { values: Record<string, Val> }).values : null
}

const cell = (s: string) => s.replaceAll('|', '\\|').replaceAll('\n', '\\n')
const row = (xs: string[]) => `| ${xs.map(cell).join(' | ')} |`

/** Which answer the engine gives. */
const follows = (o: WinOutcome, mac: Val) => (matches(o.fn, o.win, o.engine) ? 'Windows' : matches(o.fn, mac, o.engine) ? 'Mac' : 'neither')

const parted: string[] = []
const agreed: string[] = []
const counts: string[] = []
for (const set of WIN_SETS) {
  const mac = macValues(set)
  if (!mac) continue
  const { outcomes } = await runWinSet(set)
  let same = 0
  let engine = 0
  for (const o of outcomes) {
    const m = mac[o.id]
    if (!m) continue
    if (matches(o.fn, m, o.win)) {
      same++
      if (matches(o.fn, o.win, o.engine)) engine++
      else agreed.push(row([set.tag, set.name, o.id, `\`${o.formula}\``, show(o.win), show(o.engine)]))
    } else parted.push(row([set.tag, set.name, o.id, `\`${o.formula}\``, show(m), show(o.win), show(o.engine), follows(o, m)]))
  }
  counts.push(row([set.tag, set.name, String(outcomes.length), String(same), String(outcomes.filter((o) => matches(o.fn, o.win, o.engine)).length), String(engine), String(set.rejected?.length ?? 0)]))
}

console.log('## Cases compared\n')
console.log(row(['locale', 'source', 'cases', 'Windows = Mac', 'engine = Windows', 'all three agree', 'not taken by Windows']))
console.log(row(['---', '---', '---:', '---:', '---:', '---:', '---:']))
console.log(counts.join('\n'))
console.log('\n## Windows and the Mac part\n')
console.log(row(['locale', 'source', 'case', 'formula', 'Mac', 'Windows', 'engine', 'engine follows']))
console.log(row(['---', '---', '---', '---', '---', '---', '---', '---']))
console.log(parted.join('\n'))
console.log('\n## Windows and the Mac agree, the engine does not\n')
console.log(row(['locale', 'source', 'case', 'formula', 'Excel', 'engine']))
console.log(row(['---', '---', '---', '---', '---', '---']))
console.log(agreed.join('\n'))
