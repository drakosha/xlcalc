/**
 * Translation of formula text between the invariant syntax and a locale, piece by piece (`lex.ts`):
 * numbers change their decimal separator, function names, TRUE/FALSE and error literals their words,
 * separators their characters; strings, sheet names, defined names and references stay as they are, and so
 * does whitespace. The output is lexed again on the target side: a piece that would read differently there
 * (`1.5` typed in ru-RU, where the decimal separator is `,`) is a syntax error at its place in the input.
 */
import { ERROR_LITERALS, INVARIANT, lex, numberRule, tableText, type Piece, type Syntax } from './lex.js'
import type { Locale, TableItem, Translation } from './types.js'

/** Reading one syntax and translating out of it. */
interface Side {
  /** The name of the syntax this side reads, for messages. */
  name: string
  syn: Syntax
  num: RegExp
  /** Error literals this side reads. */
  errors: readonly string[]
  /** The other side's word for a function name, a boolean and an error literal of this side (upper-case keys). */
  fn: ReadonlyMap<string, string>
  word: ReadonlyMap<string, string>
  err: ReadonlyMap<string, string>
  /** The other side's specifier of a structured reference (upper-case keys of this side's, English ones too on the local side). */
  item: ReadonlyMap<string, string>
}

/** A locale's two sides: the invariant one (`toLocal` reads it) and the local one. */
export interface Sides {
  invariant: Side
  local: Side
}

type Fail = Extract<Translation, { ok: false }>

const upperKeys = (entries: [string, string][]) => new Map(entries.map(([k, v]) => [k.toUpperCase(), v]))
const byLength = (xs: string[]) => [...new Set(xs)].sort((a, b) => b.length - a.length)
const swap = (xs: [string, string][]) => xs.map(([a, b]): [string, string] => [b, a])
const same = (xs: string[]) => xs.map((x): [string, string] => [x, x])

/** Both sides of a (checked) locale. Local text also reads English function names, TRUE/FALSE and error texts. */
export function sides(loc: Locale): Sides {
  const local: Syntax = { decimal: loc.decimalSeparator, list: loc.listSeparator, arrayColumn: loc.arrayColumnSeparator, arrayRow: loc.arrayRowSeparator }
  const errors = ERROR_LITERALS.map((e): [string, string] => [e, loc.errors[e] ?? e])
  const bools: [string, string][] = [
    ['TRUE', loc.booleans.TRUE],
    ['FALSE', loc.booleans.FALSE],
  ]
  const fns = Object.entries(loc.functions)
  const items = TABLE_ITEMS.map((w): [string, string] => [w, loc.tableItems?.[w] ?? w])
  const itemKeys = (xs: [string, string][]) => new Map(xs.map(([k, v]) => [k.toLocaleUpperCase('ru-RU'), v]))
  return {
    invariant: { name: 'the invariant syntax', syn: INVARIANT, num: numberRule(INVARIANT), errors: byLength(ERROR_LITERALS), fn: upperKeys(fns), word: upperKeys(bools), err: upperKeys(errors), item: itemKeys(items) },
    local: {
      name: loc.tag,
      syn: local,
      num: numberRule(local),
      errors: byLength([...errors.map((e) => e[1]), ...ERROR_LITERALS]),
      fn: upperKeys(swap(fns)),
      // the locale's own words win over the English ones
      word: upperKeys([...same(['TRUE', 'FALSE']), ...swap(bools)]),
      err: upperKeys([...same(ERROR_LITERALS), ...swap(errors)]),
      item: itemKeys([...same(TABLE_ITEMS), ...swap(items)]),
    },
  }
}

const TABLE_ITEMS: TableItem[] = ['#All', '#Data', '#Headers', '#Totals', '#This Row']

const fail = (at: number, message: string): Fail => ({ ok: false, code: 'syntax', message, at })

/** A name in quotes because it reads as a boolean of the side (`'ИСТИНА'`), which stands for the name where it is no boolean. */
const quotedWord = (text: string, side: Side) => text.length > 2 && text.startsWith("'") && text.endsWith("'") && side.word.has(text.slice(1, -1).toUpperCase())

/** A translated text; `back` maps a position in it to the input: the same character, or the start of a translated piece. */
export interface Translated {
  ok: true
  formula: string
  back(at: number): number
}

/** Translates `src` out of the side `from` into `to`; `offset` is where `src` starts in the caller's text (after `=`). */
export function translate(src: string, from: Side, to: Side, offset = 0): Translated | Fail {
  const pieces = lex(src, from.syn, from.errors, from.num)
  if (!Array.isArray(pieces)) return fail(offset + pieces.at, pieces.message)
  // characters with a meaning on the other side: in an array constant its separators, elsewhere the list separator
  const special = [new Set([to.syn.decimal, to.syn.list]), new Set([to.syn.decimal, to.syn.arrayColumn, to.syn.arrayRow])]
  const out: Piece[] = []
  for (const p of pieces) {
    let text = p.text
    if (p.k === 'num') text = text.replace(from.syn.decimal, to.syn.decimal)
    else if (p.k === 'fn') text = from.fn.get(text.toUpperCase()) ?? text
    else if (p.k === 'err') text = from.err.get(text.toUpperCase()) ?? text
    else if (p.k === 'list') text = to.syn.list
    else if (p.k === 'tbl') text = tableText(text, from.item, from.syn.list, to.syn.list)
    else if (p.k === 'col') text = to.syn.arrayColumn
    else if (p.k === 'row') text = to.syn.arrayRow
    else if (p.k === 'word') {
      const w = from.word.get(text.toUpperCase())
      if (w === undefined && p.arr) return fail(offset + p.at, `unexpected ${JSON.stringify(text)} in an array constant`)
      // a name that reads as a boolean on the other side goes in quotes, as Excel ru-RU shows a name ИСТИНА ('ИСТИНА')
      if (w === undefined && to.word.has(text.toUpperCase())) {
        out.push({ ...p, k: 'raw', text: `'${text}'` })
        continue
      }
      text = w ?? text
    } else if (p.k === 'raw' && quotedWord(text, from) && !to.word.has(text.slice(1, -1).toUpperCase())) {
      out.push({ ...p, k: 'word', text: text.slice(1, -1) })
      continue
    } else if (special[p.arr ? 1 : 0].has(text)) return fail(offset + p.at, `unexpected ${JSON.stringify(text)}`) // `,` outside a number in ru-RU
    out.push({ ...p, text })
  }
  const formula = out.map((p) => p.text).join('')
  // reading the output on the other side must give the same pieces
  const again = lex(formula, to.syn, to.errors, to.num)
  const read = Array.isArray(again) ? again : []
  const bad = out.findIndex((p, i) => read[i]?.k !== p.k || read[i].text !== p.text)
  if (bad >= 0 || read.length !== out.length) {
    const p = pieces[bad >= 0 ? bad : out.length - 1]
    return fail(offset + p.at, `${JSON.stringify(p.text)} cannot be written in ${to.name}`)
  }
  const starts: number[] = []
  let pos = 0
  for (const p of out) {
    starts.push(pos)
    pos += p.text.length
  }
  const back = (at: number): number => {
    if (at >= formula.length) return offset + src.length
    let lo = 0
    let hi = out.length - 1
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1
      if (starts[mid] <= at) lo = mid
      else hi = mid - 1
    }
    const p = pieces[lo]
    return offset + p.at + (p.text.length === out[lo].text.length ? at - starts[lo] : 0)
  }
  return { ok: true, formula, back }
}
