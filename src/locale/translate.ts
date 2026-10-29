/**
 * Translation of formula text between the invariant syntax and a locale, piece by piece (`lex.ts`):
 * numbers change their decimal separator, function names, TRUE/FALSE and error literals their words,
 * separators their characters; strings, defined names and references stay as they are, and so does whitespace. A
 * sheet's name shown in the locale is in quotes only where Excel puts them (`'Data'!B2` is `Data!B2`); read back, it
 * stays as typed. The output is lexed again on the target side: a piece that would read differently there
 * (`1.5` typed in ru-RU, where the decimal separator is `,`) is a syntax error at its place in the input.
 */
import { quoteSheetIfNeeded } from '../a1.js'
import { wordKey } from '../collate.js'
import { LOCALE_ERRORS } from '../syntax/errors.js'
import { INVARIANT, lex, numberRule, tableText, type Piece, type Syntax } from './lex.js'
import type { Message } from './messages.js'
import { TABLE_ITEMS } from '../syntax/table.js'
import type { Locale } from './types.js'

/** Reading one syntax and translating out of it. */
interface Side {
  /** The locale's tag on its side, none on the invariant one, for messages. */
  tag?: string
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
  /** A sheet's name before a reference as this side shows it (the local side, by its booleans); none keeps the text's own. */
  sheet?: (name: string) => string
}

/** A locale's two sides: the invariant one (`toLocal` reads it) and the local one. */
export interface Sides {
  invariant: Side
  local: Side
  /** The local side reading its own error texts alone, as Excel's formula bar in that locale reads them (`refAt`). */
  own: Side
}

/** Text that does not translate: why, at its place in the caller's text. */
export interface Fail {
  ok: false
  msg: Message
  at: number
}

const upperKeys = (entries: [string, string][]) => new Map(entries.map(([k, v]) => [wordKey(k), v]))
const byLength = (xs: readonly string[]) => [...new Set(xs)].sort((a, b) => b.length - a.length)
const swap = (xs: [string, string][]) => xs.map(([a, b]): [string, string] => [b, a])
const same = (xs: readonly string[]) => xs.map((x): [string, string] => [x, x])

/** Both sides of a (checked) locale. Local text also reads English function names, TRUE/FALSE and error texts. */
export function sides(loc: Locale): Sides {
  const local: Syntax = { decimal: loc.decimalSeparator, list: loc.listSeparator, arrayColumn: loc.arrayColumnSeparator, arrayRow: loc.arrayRowSeparator }
  const errors = LOCALE_ERRORS.map((e): [string, string] => [e, loc.errors[e] ?? e])
  const bools: [string, string][] = [
    ['TRUE', loc.booleans.TRUE],
    ['FALSE', loc.booleans.FALSE],
  ]
  const fns = Object.entries(loc.functions)
  const items = TABLE_ITEMS.map((w): [string, string] => [w, loc.tableItems?.[w] ?? w])
  const side: Side = {
    tag: loc.tag,
    syn: local,
    num: numberRule(local),
    errors: byLength([...errors.map((e) => e[1]), ...LOCALE_ERRORS]),
    fn: upperKeys(swap(fns)),
    // the locale's own words win over the English ones
    word: upperKeys([...same(['TRUE', 'FALSE']), ...swap(bools)]),
    err: upperKeys([...same(LOCALE_ERRORS), ...swap(errors)]),
    item: upperKeys([...same(TABLE_ITEMS), ...swap(items)]),
    sheet: (name) => quoteSheetIfNeeded(name, loc),
  }
  return {
    invariant: { syn: INVARIANT, num: numberRule(INVARIANT), errors: byLength(LOCALE_ERRORS), fn: upperKeys(fns), word: upperKeys(bools), err: upperKeys(errors), item: upperKeys(items) },
    local: side,
    own: { ...side, errors: byLength(errors.map((e) => e[1])), err: upperKeys(swap(errors)) },
  }
}

const fail = (at: number, msg: Message): Fail => ({ ok: false, msg, at })

/** A name in quotes because it reads as a boolean of the side (`'ИСТИНА'`), which stands for the name where it is no boolean. */
const quotedWord = (text: string, side: Side) => text.length > 2 && text.startsWith("'") && text.endsWith("'") && side.word.has(wordKey(text.slice(1, -1)))

/**
 * A sheet's prefix (`'Data'!`, `Data!`, `'Jan:Mar'!`) with its names as `sheet` shows them: a reference across sheets
 * without quotes when neither end needs them, else as it is (Excel quotes the pair: `'S1:S3'!`).
 */
function sheetPrefix(text: string, sheet: (name: string) => string): string {
  const names = (text.startsWith("'") ? text.slice(1, -2).replaceAll("''", "'") : text.slice(0, -1)).split(':')
  if (names.length === 1) return `${sheet(names[0])}!`
  return names.length === 2 && names.every((n) => sheet(n) === n) ? `${names.join(':')}!` : text
}

/** Whether piece `i` is a sheet's prefix, quoted or not: not the end of a reference across sheets written without quotes (`Jan:Mar!`), nor a sheet of another book (`[1]Sheet!`). */
const isSheetPrefix = (pieces: Piece[], i: number) => {
  const { k, text } = pieces[i]
  const before = pieces[i - 1]?.text
  return k === 'raw' && text.length > 1 && text.endsWith('!') && text[0] !== '"' && before !== ':' && !before?.endsWith(']')
}

/** A translated text; `back` maps a position in it to the input: the same character, or the start of a translated piece. */
export interface Translated {
  ok: true
  formula: string
  back(at: number): number
}

/** Translates `src` out of the side `from` into `to`; `offset` is where `src` starts in the caller's text (after `=`). */
export function translate(src: string, from: Side, to: Side, offset = 0): Translated | Fail {
  const pieces = lex(src, from.syn, from.errors, from.num)
  if (!Array.isArray(pieces)) return fail(offset + pieces.at, pieces.msg)
  // characters with a meaning on the other side: in an array constant its separators, elsewhere the list separator
  const special = [new Set([to.syn.decimal, to.syn.list]), new Set([to.syn.decimal, to.syn.arrayColumn, to.syn.arrayRow])]
  const out: Piece[] = []
  for (const [i, p] of pieces.entries()) {
    let text = p.text
    if (p.k === 'num') text = text.replace(from.syn.decimal, to.syn.decimal)
    else if (p.k === 'fn') text = from.fn.get(wordKey(text)) ?? text
    else if (p.k === 'err') text = from.err.get(wordKey(text)) ?? text
    else if (p.k === 'list') text = to.syn.list
    else if (p.k === 'tbl') text = tableText(text, from.item, from.syn.list, to.syn.list)
    else if (p.k === 'col') text = to.syn.arrayColumn
    else if (p.k === 'row') text = to.syn.arrayRow
    else if (p.k === 'word') {
      const w = from.word.get(wordKey(text))
      if (w === undefined && p.arr) return fail(offset + p.at, { key: 'unexpectedInArray', text })
      // a name that reads as a boolean on the other side goes in quotes, as Excel ru-RU shows a name ИСТИНА ('ИСТИНА');
      // a table's name before its brackets does not (ИСТИНА[Qty])
      if (w === undefined && to.word.has(wordKey(text)) && pieces[i + 1]?.k !== 'tbl') {
        out.push({ ...p, k: 'raw', text: `'${text}'` })
        continue
      }
      text = w ?? text
    } else if (p.k === 'raw' && quotedWord(text, from) && !to.word.has(wordKey(text.slice(1, -1)))) {
      out.push({ ...p, k: 'word', text: text.slice(1, -1) })
      continue
    } else if (special[p.arr ? 1 : 0].has(text)) return fail(offset + p.at, { key: 'unexpectedText', text }) // `,` outside a number in ru-RU
    out.push({ ...p, text })
  }
  // reading the output on the other side must give the same pieces
  const again = lex(out.map((p) => p.text).join(''), to.syn, to.errors, to.num)
  const read = Array.isArray(again) ? again : []
  const bad = out.findIndex((p, i) => read[i]?.k !== p.k || read[i].text !== p.text)
  if (bad >= 0 || read.length !== out.length) {
    const p = pieces[bad >= 0 ? bad : out.length - 1]
    return fail(offset + p.at, { key: 'notWritable', text: p.text, ...(to.tag !== undefined ? { tag: to.tag } : {}) })
  }
  // sheets as the side shows them, after the check: `quoteSheetIfNeeded` leaves out the quotes only where the lexer reads the name without them
  const sheet = to.sheet
  if (sheet) for (let i = 0; i < out.length; i++) if (isSheetPrefix(out, i)) out[i] = { ...out[i], text: sheetPrefix(out[i].text, sheet) }
  const formula = out.map((p) => p.text).join('')
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
