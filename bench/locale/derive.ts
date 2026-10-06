/**
 * Excel's formula text in a locale (`docs/locale/excel-<tag>.json`, written by `snapshot.ts`) and the locale
 * data derived from it: separators from the international settings, TRUE/FALSE and error literals from
 * their probes, and function names by pairing the calls of each English formula with those of its local
 * text, in order. The scanner here is independent of the library's lexer, so the tests compare two readings.
 */
import { readFileSync } from 'node:fs'

export interface Snapshot {
  meta: { excel: string; mode: string; date: string; seconds: number }
  /** `Application.International` by setting. */
  international: Record<string, string | number | boolean>
  /** Functions of the reference list (`bench/coverage/excel-functions.json`). */
  functions: string[]
  /** Probes Excel does not load. */
  rejected: string[]
  /** By probe id: Excel's English formula text and its local text, without `=`. */
  formulas: Record<string, [string, string]>
}

export const snapshotFile = (tag: string) => new URL(`../../docs/locale/excel-${tag}.json`, import.meta.url)

export const loadSnapshot = (tag: string): Snapshot => JSON.parse(readFileSync(snapshotFile(tag), 'utf8')) as Snapshot

/** JSON with one formula per line, so a new snapshot diffs line by line. */
export function serialize(s: Snapshot): string {
  const head = JSON.stringify({ meta: s.meta, international: s.international, functions: s.functions, rejected: s.rejected }, null, 1).slice(0, -2)
  const lines = Object.entries(s.formulas).map(([id, pair]) => ` ${JSON.stringify(id)}: ${JSON.stringify(pair)}`)
  return `${head},\n "formulas": {\n${lines.join(',\n')}\n }\n}\n`
}

export interface Derived {
  decimalSeparator: string
  listSeparator: string
  arrayColumnSeparator: string
  arrayRowSeparator: string
  booleans: { TRUE: string; FALSE: string }
  /** Error literals Excel loads, by invariant text. */
  errors: Record<string, string>
  /** Every function of the list Excel knows, by English name (the same name included). */
  functions: Record<string, string>
  /** Functions of the list this Excel does not know: their local text keeps the `_xlfn.` prefix. */
  unknown: string[]
  /** Anything that does not fit: calls that do not pair, two local names for one function, functions never seen. */
  problems: string[]
}

/** Names of the functions a formula calls, in order: strings and quoted sheet names skipped. */
export function calls(formula: string): string[] {
  const bare = formula.replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'/g, '""')
  return [...bare.matchAll(/(?<![\p{L}\p{N}_.\\])([\p{L}_\\][\p{L}\p{N}_.]*)\s*\(/gu)].map((m) => m[1])
}

const PREFIX = /^_xlfn\.(?:_xlws\.)?/

export function derive(s: Snapshot): Derived {
  const intl = (k: string) => String(s.international[k])
  const problems: string[] = []
  const local = new Map<string, string>()
  const unknown = new Set<string>()
  for (const [id, [en, loc]] of Object.entries(s.formulas)) {
    const a = calls(en)
    const b = calls(loc)
    if (a.length !== b.length) {
      problems.push(`${id}: ${a.length} calls in ${en}, ${b.length} in ${loc}`)
      continue
    }
    a.forEach((name, i) => {
      const fn = name.replace(PREFIX, '').toUpperCase()
      if (PREFIX.test(b[i])) return void unknown.add(fn)
      const was = local.get(fn)
      if (was !== undefined && was !== b[i]) problems.push(`${id}: ${fn} is ${b[i]} here, ${was} before`)
      local.set(fn, b[i])
    })
  }
  const functions: Record<string, string> = {}
  for (const f of s.functions) {
    const l = local.get(f)
    if (l !== undefined) functions[f] = l
    else if (!unknown.has(f)) problems.push(`${f}: no call seen`)
  }
  const errors: Record<string, string> = {}
  for (const [id, [en, loc]] of Object.entries(s.formulas)) if (id.startsWith('err:')) errors[en] = loc
  const lit = (id: string) => s.formulas[id]?.[1] ?? ''
  return {
    decimalSeparator: intl('decimal separator'),
    listSeparator: intl('list separator'),
    arrayColumnSeparator: intl('column separator'),
    arrayRowSeparator: intl('row separator'),
    booleans: { TRUE: lit('lit:1'), FALSE: lit('lit:2') },
    errors,
    functions,
    unknown: s.functions.filter((f) => unknown.has(f) && !local.has(f)),
    problems,
  }
}
