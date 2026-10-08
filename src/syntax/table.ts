/**
 * Structured references to tables (`Sales[Qty]`, `Sales[[#Headers],[#Data],[Qty]:[Price]]`, `[@Qty]`): reading the
 * part in brackets and printing it as an xlsx file stores it. `'` takes the next character as it is (`[`, `]`, `#`,
 * `'`, and `@` before a column name); `@` before a column is the formula's row, `[#This Row]` in stored text.
 */
import { bad } from './ast.js'

/** A table's rows a reference takes: all of them, the data, the header row, the totals row, the formula's row. */
export type TableArea = 'all' | 'data' | 'headers' | 'totals' | 'row'

/** What a structured reference names in its brackets: areas in Excel's order, and one column or a span of two. */
export interface TableSpec {
  areas: TableArea[]
  cols?: [string] | [string, string]
}

/** The specifiers' words in stored text; the reader takes them in any case. */
export const AREA_WORDS: Record<TableArea, string> = { all: '#All', headers: '#Headers', data: '#Data', totals: '#Totals', row: '#This Row' }

const ORDER: TableArea[] = ['all', 'headers', 'data', 'totals', 'row']

/** The sets of areas Excel takes (by their order): one alone, headers with data, data with totals. */
const ALLOWED = new Set(['', 'all', 'headers', 'data', 'totals', 'row', 'headers,data', 'data,totals'])

/** An area by its word in any case, `words` (a locale's) beside the stored ones. */
export function areaOf(word: string, words: Record<TableArea, string> = AREA_WORDS): TableArea | undefined {
  const w = word.toLocaleUpperCase('ru-RU')
  for (const k of ORDER) if (AREA_WORDS[k].toUpperCase() === w || words[k].toLocaleUpperCase('ru-RU') === w) return k
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

/** How a structured reference is printed: a locale's specifier words and list separator, the file's escapes. */
export interface SpecStyle {
  words?: Record<TableArea, string>
  list?: string
  file?: boolean
}

/** The text between a structured reference's brackets as stored: `Qty`, `[Qty]:[Price]`, `#All`, `[#This Row],[Qty]`, empty for the data. */
export function specText(spec: TableSpec, style: SpecStyle = {}): string {
  const { words = AREA_WORDS, list = ',', file = false } = style
  const esc = (c: string) => escapeColumn(c, file)
  const cols = spec.cols ? spec.cols.map((c) => `[${esc(c)}]`).join(':') : ''
  if (!spec.areas.length) return spec.cols?.length === 1 ? esc(spec.cols[0]) : cols
  if (spec.areas.length === 1 && !spec.cols) return words[spec.areas[0]]
  return [...spec.areas.map((a) => `[${words[a]}]`), ...(cols ? [cols] : [])].join(list)
}

/**
 * Reads the bracketed part of a structured reference starting at the `[` at `i` of `src`; `at` is the reference's
 * position for messages. Spaces may stand inside the outer brackets around a list of items and after its commas,
 * not around `:`; a column written alone (`Sales[Qty]`) is all the text to its `]`, spaces included.
 */
export function readSpec(src: string, i: number, at: number, words?: Record<TableArea, string>, list = ','): { spec: TableSpec; end: number } {
  let j = i + 1
  const fail = (message: string): never => bad(at, message)
  const skip = () => {
    while (src[j] === ' ') j++
  }
  // a name up to the first `]` not escaped by `'`
  const name = (): string => {
    let out = ''
    for (;;) {
      const ch = src[j]
      if (ch === undefined) fail('unterminated structured reference')
      if (ch === "'") {
        if (src[j + 1] === undefined) fail('unterminated structured reference')
        out += src[j + 1]
        j += 2
        continue
      }
      if (ch === ']') return out
      if (ch === '[') fail('unexpected [ in a column name (write it as \'[)')
      out += ch
      j++
    }
  }
  const area = (): TableArea => {
    const w = `#${name()}`
    return areaOf(w, words) ?? fail(`unknown table specifier ${w}`)
  }
  // `[Col]` or `[Col]:[Col2]` at j (on its `[`)
  const column = (): [string] | [string, string] => {
    j++
    const a = name()
    j++
    if (src[j] !== ':') return [a]
    if (src[j + 1] !== '[') fail('expected [ after : in a structured reference')
    j += 2
    const b = name()
    j++
    return [a, b]
  }
  const areas = new Set<TableArea>()
  let cols: [string] | [string, string] | undefined
  const add = (a: TableArea) => {
    if (areas.has(a)) fail(`${AREA_WORDS[a]} twice in a structured reference`)
    areas.add(a)
  }
  let k = j
  while (src[k] === ' ') k++
  if (src[k] === ']') j = k
  else if (src[k] === '[') {
    j = k
    for (;;) {
      if (src[j] !== '[') fail('expected [ in a structured reference')
      if (src[j + 1] === '#') {
        j += 2
        add(area())
        j++
      } else {
        if (cols) fail('two columns in a structured reference (write a span as [A]:[B])')
        cols = column()
      }
      skip()
      if (src[j] !== list) break
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
  if (src[j] !== ']') fail('expected ] to close a structured reference')
  const sorted = ORDER.filter((a) => areas.has(a))
  if (!ALLOWED.has(sorted.join(','))) fail(`a structured reference cannot take ${sorted.map((a) => AREA_WORDS[a]).join(' and ')} together`)
  return { spec: cols ? { areas: sorted, cols } : { areas: sorted }, end: j + 1 }
}

/** Characters of a column name that Excel shows in brackets after `@` (`[@[Unit Price]]`). */
const SPECIAL = /[\s,:.[\]#'"{}$^&*+=\-<>/@]/

/**
 * The text between a structured reference's brackets as Excel shows it: the formula's row as `@` (`@Qty`, `@[Unit
 * Price]`, `@[Qty]:[Price]`, `@` alone), the rest as stored.
 */
export function shownSpec(spec: TableSpec, style: SpecStyle = {}): string {
  if (spec.areas.length !== 1 || spec.areas[0] !== 'row') return specText(spec, style)
  if (!spec.cols) return '@'
  const [a, b] = spec.cols
  if (b !== undefined) return `@[${escapeColumn(a)}]:[${escapeColumn(b)}]`
  return SPECIAL.test(a) ? `@[${escapeColumn(a)}]` : `@${escapeColumn(a)}`
}

/** Whether Excel shows a reference in its own table without the table's name: one column or a span, of the data or the formula's row. */
export const unnamedInTable = (spec: TableSpec) => !!spec.cols && (spec.areas.length === 0 || (spec.areas.length === 1 && spec.areas[0] === 'row'))
