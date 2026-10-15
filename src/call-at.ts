/**
 * The call a caret stands in, in formula text as it is typed (`callAt`), for the tip of a function's arguments that a
 * grid shows while a formula is typed: the innermost function whose brackets hold the caret and the argument it is in.
 * The text is read as `refAt` reads it, character by character, with no parser: a string, a quoted sheet's name and a
 * table's brackets hold no separator or bracket, an array constant (`{1,2;3}`) and a group's brackets (`(A1,B1)`) hold
 * separators of their own, so half typed or broken text still gives its call. The package's own rule, not measured in
 * Excel.
 */
import type { CallAt, ScanOptions } from './contract.js'
import { checkNumber, checkString, object, onlyKeys } from './input.js'
import { checkLocaleArg } from './locale/load.js'
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
 * The innermost call whose brackets hold `caret` (the end of the text by default) of formula text `text`, and the argument
 * it is in; null outside any call, and for text that is no formula (not begun by `=`, `+` or `-`). `opts.locale`: the
 * locale the text is typed in (its list separator and names); without it the text is en-US's.
 */
export function callAt(text: string, caret?: number, opts?: ScanOptions): CallAt | null {
  const t = checkString(text, 'text')
  if (caret !== undefined) checkNumber(caret, 'caret', 0, t.length, true)
  const at = caret ?? t.length
  if (opts !== undefined) onlyKeys(object(opts, 'opts', "options such as { locale: 'ru-RU' }"), 'opts', ['locale'])
  const locale = opts?.locale === undefined ? undefined : checkLocaleArg(opts.locale, 'opts.locale')
  if (!/^[=+-]/.test(t)) return null
  const sep = locale?.listSeparator ?? ','
  const stack: Frame[] = []
  let quote: string | null = null
  let table = 0
  for (let i = 1; i < at; i++) {
    const c = t[i]
    if (quote) {
      if (c === quote) quote = null
    } else if (c === '"' || c === "'") quote = c
    else if (c === '[') table++
    else if (c === ']') table = Math.max(0, table - 1)
    else if (table) continue
    else if (c === '{') stack.push({ kind: 'array' })
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
  }
  const call = stack.findLast((f) => f.kind === 'call') as Extract<Frame, { kind: 'call' }> | undefined
  if (!call) return null
  return { name: call.name, fn: englishName(call.name, locale), index: call.index, open: call.open, close: closeOf(t, call.open) }
}

/** Where the bracket opened at `open` closes, read as `callAt` reads; null while it is open. */
function closeOf(t: string, open: number): number | null {
  let quote: string | null = null
  let table = 0
  let depth = 0
  for (let i = open; i < t.length; i++) {
    const c = t[i]
    if (quote) {
      if (c === quote) quote = null
    } else if (c === '"' || c === "'") quote = c
    else if (c === '[') table++
    else if (c === ']') table = Math.max(0, table - 1)
    else if (table) continue
    else if (c === '(') depth++
    else if (c === ')' && --depth === 0) return i
  }
  return null
}
