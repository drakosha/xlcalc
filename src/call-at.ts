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
import type { Locale } from './locale/types.js'

/** A character of a function's name: a letter of any script, a digit, `.` or `_` (`СТАНДОТКЛОН.В`, `LOG10`). */
const NAME_CHAR = /[\p{L}\p{N}._]/u

/** A bracket open at the caret: a call's with its name, its index so far and where it opened, or a group's or an array's. */
type Frame = { kind: 'call'; name: string; index: number; open: number } | { kind: 'group' | 'array' }

/** A function's English name: the locale's name translated, in upper case; another name upper-cased as it is. */
function englishName(name: string, locale: Locale | undefined): string {
  const upper = name.toUpperCase()
  if (!locale) return upper
  const fns = locale.functions
  return Object.keys(fns).find((k) => fns[k].toUpperCase() === upper) ?? upper
}

/**
 * The characters of `t` from `from` up to `to` outside strings, quoted sheets' names and tables' brackets, each given to
 * `f` with its place; `f` answering true stops the walk. What is open where it stopped: a string's or a quoted sheet's
 * quote, and tables' brackets.
 */
export function walk(t: string, from: number, to: number, f: (c: string, i: number) => boolean | void): { quote: string | null; table: number } {
  let quote: string | null = null
  let table = 0
  for (let i = from; i < to; i++) {
    const c = t[i]
    if (quote) {
      if (c === quote) quote = null
    } else if (c === '"' || c === "'") quote = c
    else if (c === '[') table++
    else if (c === ']') table = Math.max(0, table - 1)
    else if (!table && f(c, i)) break
  }
  return { quote, table }
}

/**
 * The innermost call whose brackets hold `caret` (the end of the text by default) of formula text `text`, and the argument
 * it is in; null outside any call, and for text that is no formula (not begun by `=`, `+` or `-`). `opts.locale`: the
 * locale the text is typed in (its list separator and names); without it the text is en-US's.
 */
export function callAt(text: string, caret?: number, opts?: ScanOptions): CallAt | null {
  const { t, at, locale } = typedArgs(text, caret, opts)
  if (!/^[=+-]/.test(t)) return null
  const sep = locale?.listSeparator ?? ','
  const stack: Frame[] = []
  walk(t, 1, at, (c, i) => {
    if (c === '{') stack.push({ kind: 'array' })
    else if (c === '}') {
      if (stack.at(-1)?.kind === 'array') stack.pop()
    } else if (c === '(') {
      let from = i
      while (from > 1 && NAME_CHAR.test(t[from - 1])) from--
      stack.push(from < i ? { kind: 'call', name: t.slice(from, i), index: 0, open: i } : { kind: 'group' })
    } else if (c === ')') {
      if (stack.at(-1)?.kind !== 'array') stack.pop()
    } else if (c === sep) {
      const top = stack.at(-1)
      if (top?.kind === 'call') top.index++
    }
  })
  const call = stack.findLast((f) => f.kind === 'call') as Extract<Frame, { kind: 'call' }> | undefined
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
