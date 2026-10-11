/**
 * Structured references to tables (`Sales[Qty]`, `Sales[[#Headers],[#Data],[Qty]:[Price]]`, `[@Qty]`): reading the
 * part in brackets and printing it as an xlsx file stores it. `'` takes the next character as it is (`[`, `]`, `#`,
 * `'`, and `@` before a column name); `@` before a column is the formula's row, `[#This Row]` in stored text.
 */
import { caseKey } from '../collate.js'
import type { Message } from '../locale/messages.js'
import { bad } from './ast.js'

/** A table's rows a reference takes: all of them, the data, the header row, the totals row, the formula's row. */
export type TableArea = 'all' | 'data' | 'headers' | 'totals' | 'row'

/** What a structured reference names in its brackets: areas in Excel's order, and one column or a span of two. */
export interface TableSpec {
  areas: TableArea[]
  cols?: [string] | [string, string]
}

/** The specifiers' words in stored text, a locale's `tableItems` keyed by them; the reader takes them in any case. */
export const AREA_WORDS = { all: '#All', data: '#Data', headers: '#Headers', totals: '#Totals', row: '#This Row' } as const satisfies Record<TableArea, `#${string}`>

/** A specifier of structured references in stored text. */
export type TableItem = (typeof AREA_WORDS)[TableArea]

/** The specifiers in the order a locale's check lists them (`AREA_WORDS`' own). */
export const TABLE_ITEMS: readonly TableItem[] = Object.values(AREA_WORDS)

const ORDER: TableArea[] = ['all', 'headers', 'data', 'totals', 'row']

/** The sets of areas Excel takes (by their order): one alone, headers with data, data with totals. */
const ALLOWED = new Set(['', 'all', 'headers', 'data', 'totals', 'row', 'headers,data', 'data,totals'])

/** An area by its word in any case. */
function areaOf(word: string): TableArea | undefined {
  const w = caseKey(word)
  for (const k of ORDER) if (caseKey(AREA_WORDS[k]) === w) return k
  return undefined
}

/**
 * A column name with `'` before the characters it would otherwise mean: `[`, `]`, `#`, `'`, and a leading `@` (but in
 * an xlsx file, `file`, where `@` means nothing).
 */
export const escapeColumn = (name: string, file = false) => {
  const out = name.replace(/['[\]#]/g, "'$&")
  return file ? out : out.replace(/^@/, "'@")
}

/**
 * The text between a structured reference's brackets as stored: `Qty`, `[Qty]:[Price]`, `#All`, `[#This Row],[Qty]`,
 * empty for the data; `file`, with the escapes of an xlsx file.
 */
export function specText(spec: TableSpec, { file = false }: { file?: boolean } = {}): string {
  const esc = (c: string) => escapeColumn(c, file)
  const cols = spec.cols ? spec.cols.map((c) => `[${esc(c)}]`).join(':') : ''
  if (!spec.areas.length) return spec.cols?.length === 1 ? esc(spec.cols[0]) : cols
  if (spec.areas.length === 1 && !spec.cols) return AREA_WORDS[spec.areas[0]]
  return [...spec.areas.map((a) => `[${AREA_WORDS[a]}]`), ...(cols ? [cols] : [])].join(',')
}

/**
 * Reads the bracketed part of a structured reference starting at the `[` at `i` of `src`; `at` is the reference's
 * position for messages. Spaces may stand inside the outer brackets around a list of items and after its commas,
 * not around `:`; a column written alone (`Sales[Qty]`) is all the text to its `]`, spaces included.
 */
export function readSpec(src: string, i: number, at: number): { spec: TableSpec; end: number } {
  let j = i + 1
  const fail = (msg: Message): never => bad(at, msg)
  const skip = () => {
    while (src[j] === ' ') j++
  }
  // a name up to the first `]` not escaped by `'`
  const name = (): string => {
    let out = ''
    for (;;) {
      const ch = src[j]
      if (ch === undefined) fail({ key: 'specUnterminated' })
      if (ch === "'") {
        if (src[j + 1] === undefined) fail({ key: 'specUnterminated' })
        out += src[j + 1]
        j += 2
        continue
      }
      if (ch === ']') return out
      if (ch === '[') fail({ key: 'specBracketInColumn' })
      out += ch
      j++
    }
  }
  const area = (): TableArea => {
    const w = `#${name()}`
    return areaOf(w) ?? fail({ key: 'specUnknown', word: w })
  }
  // `[Col]` or `[Col]:[Col2]` at j (on its `[`)
  const column = (): [string] | [string, string] => {
    j++
    const a = name()
    j++
    if (src[j] !== ':') return [a]
    if (src[j + 1] !== '[') fail({ key: 'specSpan' })
    j += 2
    const b = name()
    j++
    return [a, b]
  }
  const areas = new Set<TableArea>()
  let cols: [string] | [string, string] | undefined
  const add = (a: TableArea) => {
    if (areas.has(a)) fail({ key: 'specTwice', word: AREA_WORDS[a] })
    areas.add(a)
  }
  let k = j
  while (src[k] === ' ') k++
  if (src[k] === ']') j = k
  else if (src[k] === '[') {
    j = k
    for (;;) {
      if (src[j] !== '[') fail({ key: 'specOpen' })
      if (src[j + 1] === '#') {
        j += 2
        add(area())
        j++
      } else {
        if (cols) fail({ key: 'specTwoColumns' })
        cols = column()
      }
      skip()
      if (src[j] !== ',') break
      j++
      skip()
    }
    skip()
  } else if (src[j] === '#') {
    j++
    add(area())
  } else if (src[j] === '@') {
    j++
    add('row')
    if (src[j] === '[') cols = column()
    else if (src[j] !== ']') cols = [name()]
  } else cols = [name()]
  if (src[j] !== ']') fail({ key: 'specClose' })
  const sorted = ORDER.filter((a) => areas.has(a))
  if (!ALLOWED.has(sorted.join(','))) fail({ key: 'specTogether', words: sorted.map((a) => AREA_WORDS[a]) })
  return { spec: cols ? { areas: sorted, cols } : { areas: sorted }, end: j + 1 }
}

/** Characters of a column name that Excel shows in brackets after `@` (`[@[Unit Price]]`), with the list separator of the formula's locale. */
const SPECIAL = /[\s:.[\]#'"{}$^&*+=\-<>/@]/

/**
 * The text between a structured reference's brackets as Excel shows it: the formula's row as `@` (`@Qty`, `@[Unit
 * Price]`, `@[Qty]:[Price]`, `@` alone), the rest as stored. `list`, the list separator of the locale it is shown in,
 * puts a name in brackets after `@` too (`@[1,5]` in en-US, `@1,5` in ru-RU: the set `tables/headers-ru-held`).
 */
export function shownSpec(spec: TableSpec, list = ','): string {
  if (spec.areas.length !== 1 || spec.areas[0] !== 'row') return specText(spec)
  if (!spec.cols) return '@'
  const [a, b] = spec.cols
  if (b !== undefined) return `@[${escapeColumn(a)}]:[${escapeColumn(b)}]`
  return SPECIAL.test(a) || a.includes(list) ? `@[${escapeColumn(a)}]` : `@${escapeColumn(a)}`
}

/** Whether Excel shows a reference in its own table without the table's name: one column or a span, of the data or the formula's row. */
export const unnamedInTable = (spec: TableSpec) => !!spec.cols && (spec.areas.length === 0 || (spec.areas.length === 1 && spec.areas[0] === 'row'))
