/**
 * Excel for Windows against Excel for Mac and the engine, case by case, over the formula answers stored in
 * `test/excel/win/` (`npx tsx bench/windows/compare.ts > bench/windows/differences.md`): a table of the cases where
 * Windows and the Mac part, with the engine's answer, which of the two it follows and the class (a platform difference
 * with its reason in `WIN_REASONS`, or noise), then the cases where both agree and the engine does not. Numbers compare
 * as the tests compare them.
 */
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { Val } from '../coverage/formulajs.js'
import { localeDir, measurementFile } from '../coverage/locales.js'
import { KNOWN, matches, show } from '../../test/excel-cases.js'
import { KNOWN_BY_LOCALE } from '../../test/excel-locale-cases.js'
import { runWinSet, WIN_KNOWN, WIN_REASONS, WIN_SETS, type WinOutcome, type WinSet } from '../../test/excel-win-cases.js'

const ROOT = join(fileURLToPath(new URL('.', import.meta.url)), '../..')

/** The Mac's answers to the same cases, from the file the Mac harness stored. */
function macValues(set: WinSet): Record<string, Val> | null {
  const rel = set.name === 'measurement' ? `coverage/${measurementFile(set.tag)}` : set.name === 'probes' ? 'probes.json' : `cases/${localeDir(set.tag)}${set.name}.json`
  const file = join(ROOT, 'test/excel', rel)
  return existsSync(file) ? (JSON.parse(readFileSync(file, 'utf8')) as { values: Record<string, Val> }).values : null
}

const cell = (s: string) => s.replaceAll('|', '\\|').replaceAll('\n', '\\n')
const row = (xs: string[]) => `| ${xs.map(cell).join(' | ')} |`

/** Which answer the engine gives. */
const follows = (o: WinOutcome, mac: Val) => (matches(o.fn, o.win, o.engine) ? 'Windows' : matches(o.fn, mac, o.engine) ? 'Mac' : 'neither')

/**
 * The class of a case where Windows and the Mac part: an Excel platform difference (the reason pinned in WIN_KNOWN, or
 * the engine follows Windows, as for CHAR and CODE), or noise where the Mac's own answer is known not to repeat.
 */
function classOf(set: WinSet, o: WinOutcome): string {
  const mac = (set.tag === 'en-US' ? KNOWN : KNOWN_BY_LOCALE[set.tag])?.[`${set.name}:${o.id}`]
  if (mac?.reason === 'excelContext') return 'noise (the Mac\'s answer does not repeat)'
  const win = WIN_KNOWN[o.key]?.reason
  if (win && win in WIN_REASONS) return `platform: ${win}`
  return matches(o.fn, o.win, o.engine) ? 'platform: the engine follows Windows' : 'unclassified'
}

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
      else agreed.push(row([set.tag, set.name, o.id, `\`${o.formula}\``, show(o.win), show(o.engine), WIN_KNOWN[o.key]?.reason ?? '']))
    } else parted.push(row([set.tag, set.name, o.id, o.fn ?? o.group, `\`${o.formula}\``, show(m), show(o.win), show(o.engine), follows(o, m), classOf(set, o)]))
  }
  counts.push(row([set.tag, set.name, String(outcomes.length), String(same), String(outcomes.filter((o) => matches(o.fn, o.win, o.engine)).length), String(engine), String(set.rejected?.length ?? 0)]))
}

console.log('# Excel for Windows against Excel for Mac and the engine\n')
console.log('Written by `npx tsx bench/windows/compare.ts` from the formula answers in `test/excel/win/` and `test/excel/`. The class of a parted case: a platform difference (the reason in `WIN_REASONS` of `test/excel-win-cases.ts`, or the engine follows Windows) or noise.\n')
console.log('## Cases compared\n')
console.log(row(['locale', 'source', 'cases', 'Windows = Mac', 'engine = Windows', 'all three agree', 'not taken by Windows']))
console.log(row(['---', '---', '---:', '---:', '---:', '---:', '---:']))
console.log(counts.join('\n'))
console.log('\n## Windows and the Mac part\n')
console.log(row(['locale', 'source', 'case', 'function', 'formula', 'Mac', 'Windows', 'engine', 'engine follows', 'class']))
console.log(row(['---', '---', '---', '---', '---', '---', '---', '---', '---', '---']))
console.log(parted.join('\n'))
console.log('\n## Windows and the Mac agree, the engine does not\n')
console.log(row(['locale', 'source', 'case', 'formula', 'Excel', 'engine', 'reason (the Mac\'s, test/excel-cases.ts)']))
console.log(row(['---', '---', '---', '---', '---', '---', '---']))
console.log(agreed.join('\n'))
