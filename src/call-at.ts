/**
 * The call a caret stands in, in formula text as it is typed (`callAt`), for the tip of a function's arguments that a
 * grid shows while a formula is typed: the innermost function whose brackets hold the caret and the argument it is in.
 * The text is read as `refAt` reads it, character by character, with no parser: a string, a quoted sheet's name and a
 * table's brackets hold no separator or bracket, an array constant (`{1,2;3}`) and a group's brackets (`(A1,B1)`) hold
 * separators of their own, so half typed or broken text still gives its call. The package's own rule, not measured in
 * Excel. And the name of a function being typed at a caret (`nameAt`), for Formula AutoComplete.
 */
import type { CallAt, NameAt, ScanOptions } from './contract.js'
import { typedArgs } from './point.js'
import { NAME_CHAR, openBrackets, walk, type OpenBracket } from './syntax/scan.js'
import type { Locale } from './locale/types.js'

/** A function's English name: the locale's name translated, in upper case; another name upper-cased as it is. */
export function englishName(name: string, locale: Locale | undefined): string {
  const upper = name.toUpperCase()
  if (!locale) return upper
  const fns = locale.functions
  return Object.keys(fns).find((k) => fns[k].toUpperCase() === upper) ?? upper
}

/**
 * The innermost call whose brackets hold `caret` (the end of the text by default) of formula text `text`, and the argument
 * it is in; null outside any call, and for text that is no formula (not begun by `=`, `+` or `-`). `opts.locale`: the
 * locale the text is typed in (its list separator and names); without it the text is en-US's.
 */
export function callAt(text: string, caret?: number, opts?: ScanOptions): CallAt | null {
  const { t, at, locale } = typedArgs(text, caret, opts)
  if (!/^[=+-]/.test(t)) return null
  const stack = openBrackets(t, at, locale?.listSeparator).open
  const call = stack.findLast((f) => f.kind === 'call') as Extract<OpenBracket, { kind: 'call' }> | undefined
  if (!call) return null
  let close: number | null = null
  let depth = 0
  walk(t, call.open, t.length, (c, i) => {
    if (c === '(') depth++
    else if (c === ')' && --depth === 0) {
      close = i
      return true
    }
  })
  return { name: call.name, fn: englishName(call.name, locale), index: call.index, open: call.open, close }
}

/** What may stand right before a function's name, beside the list separator: the formula's first character, an operator, `(`, a space or a line break. */
const BEFORE_NAME = /[=+\-*/^&<>@(\s]/

/**
 * The name of a function being typed at `caret` (the end of the text by default) of formula text `text`, for the list
 * of Formula AutoComplete: the characters of a name around the caret, begun by a letter or `_`, where an operand may
 * begin (after the formula's first character, an operator, `(`, the locale's list separator, a space); null where no
 * function can stand: no name at the caret, in a string, a quoted sheet's name or a table's brackets, after a sheet's
 * `!` or a `$`, and for text that is no formula. `opts.locale`: the locale the text is typed in (its list separator);
 * without it the text is en-US's.
 */
export function nameAt(text: string, caret?: number, opts?: ScanOptions): NameAt | null {
  const { t, at, locale } = typedArgs(text, caret, opts)
  if (!/^[=+-]/.test(t)) return null
  let from = at
  while (from > 1 && NAME_CHAR.test(t[from - 1])) from--
  if (from === at || !/[\p{L}_]/u.test(t[from])) return null
  const before = t[from - 1]
  if (!BEFORE_NAME.test(before) && before !== (locale?.listSeparator ?? ',')) return null
  const { quote, table } = walk(t, 1, from, () => {})
  if (quote || table) return null
  let to = at
  while (to < t.length && NAME_CHAR.test(t[to])) to++
  return { from, to, typed: t.slice(from, at) }
}
