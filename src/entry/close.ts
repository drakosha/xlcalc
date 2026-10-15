/**
 * A typed formula's brackets, quotes and array braces left open at its end, closed as Excel closes them on Return
 * (`bench/coverage/entry/close-held.ts`, `close-ru-held.ts`): Excel enters a formula begun by a function, with no `)`
 * typed and only functions' brackets left open, closed without a word (`=SUM(1,2` is `=SUM(1,2)`); any other formula
 * left open it offers to correct, in its window, with what is left open closed (`=(1+2` to `=(1+2)`).
 */

/** A bracket or brace open at the end of the text: a function's, a group's, an array constant's. */
type Open = 'call' | 'group' | 'array'

/** A function's name before its bracket: a letter of any script or `_`, then letters, digits, `.` and `_`. */
const NAME = /[\p{L}_][\p{L}\p{N}._]*$/u
const BEGINS_WITH_CALL = /^[\p{L}_][\p{L}\p{N}._]*\(/u

/**
 * The text of a typed formula (begun by `=`, `+` or `-`) with what it leaves open closed at its end, whether Excel
 * closes it without a word, and the names of the functions it calls; null when nothing is left open, or a quoted sheet's
 * name or a table's brackets are.
 */
export function closeTyped(text: string): { text: string; silent: boolean; calls: string[] } | null {
  let quote: string | null = null
  let table = 0
  let closed = false
  const open: Open[] = []
  const calls: string[] = []
  for (let i = 1; i < text.length; i++) {
    const c = text[i]
    if (quote) {
      if (c === quote) quote = null
    } else if (c === '"' || c === "'") quote = c
    else if (c === '[') table++
    else if (c === ']') table = Math.max(0, table - 1)
    else if (table) continue
    else if (c === '{') open.push('array')
    else if (c === '}') {
      if (open.at(-1) === 'array') open.pop()
    } else if (c === '(') {
      const name = NAME.exec(text.slice(1, i))?.[0]
      if (name) calls.push(name)
      open.push(name ? 'call' : 'group')
    } else if (c === ')') {
      closed = true
      if (open.at(-1) !== 'array') open.pop()
    }
  }
  if (table || quote === "'" || (!quote && !open.length)) return null
  const tail = (quote ? '"' : '') + open.map((o) => (o === 'array' ? '}' : ')')).reverse().join('')
  return { text: text + tail, silent: !quote && !closed && open.every((o) => o === 'call') && BEGINS_WITH_CALL.test(text.slice(1)), calls }
}
