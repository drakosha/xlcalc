/**
 * The call a caret stands in, in formula text as it is typed (`callAt`), for the tip of a function's arguments that a
 * grid shows while a formula is typed: the innermost function whose brackets hold the caret and the argument it is in.
 * The text is read as `refAt` reads it, character by character, with no parser: a string, a quoted sheet's name and a
 * table's brackets hold no separator or bracket, an array constant (`{1,2;3}`) and a group's brackets (`(A1,B1)`) hold
 * separators of their own, so half typed or broken text still gives its call. The package's own rule, not measured in
 * Excel.
 */
import type { CallAt, ScanOptions } from './contract.js'
import { typedArgs } from './point.js'
import { openBrackets, walk, type OpenBracket } from './syntax/scan.js'
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
