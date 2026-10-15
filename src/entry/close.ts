/**
 * A typed formula's brackets, quotes and array braces left open at its end, closed as Excel closes them on Return
 * (`bench/coverage/entry/close-*.ts`): Excel enters a formula begun by a function, a sign before it or not, with no `)`
 * typed and only that function's bracket left open, closed without a word (`=SUM(1,2` is `=SUM(1,2)`, `=-SUM(1` is
 * `=-SUM(1)`, `= SUM(1` is `= SUM(1)`); any other formula left open it offers to correct, in its window, with what is
 * left open closed and the spaces after `=` dropped (`=(1+2` to `=(1+2)`, `=SUM(MAX(1` to `=SUM(MAX(1))`, `= (1+2` to
 * `=(1+2)`), but an array constant left open only alone in a formula begun by an array constant (`={1,2` to `={1,2}`,
 * `={1,2}+{3` to `={1,2}+{3}`; not `=SUM({1,2` nor `=1+{1,2`).
 */
import { walk } from '../call-at.js'

/** A bracket or brace open at the end of the text: a function's, a group's, an array constant's. */
type Open = 'call' | 'group' | 'array'

/** A function's name before its bracket: a letter of any script or `_`, then letters, digits, `.` and `_`. */
const NAME = /[\p{L}_][\p{L}\p{N}._]*$/u
/** A formula's text after its first character that begins with a function's bracket, signs and spaces before it or not. */
const BEGINS_WITH_CALL = /^[\s+-]*[\p{L}_][\p{L}\p{N}._]*\(/u

/**
 * The text of a typed formula (begun by `=`, `+` or `-`) with what it leaves open closed at its end, the text Excel
 * offers as its correction, whether it closes it without a word, whether it may offer the correction, and the names of
 * the functions it calls; null when nothing is left open, or a quoted sheet's name or a table's brackets are.
 */
export function closeTyped(text: string): { closed: string; correction: string; silent: boolean; offered: boolean; calls: string[] } | null {
  let shut = false
  const open: Open[] = []
  const calls: string[] = []
  const { quote, table } = walk(text, 1, text.length, (c, i) => {
    if (c === '{') open.push('array')
    else if (c === '}') {
      if (open.at(-1) === 'array') open.pop()
    } else if (c === '(') {
      const name = NAME.exec(text.slice(1, i))?.[0]
      if (name) calls.push(name)
      open.push(name ? 'call' : 'group')
    } else if (c === ')') {
      shut = true
      if (open.at(-1) !== 'array') open.pop()
    }
  })
  if (table || quote === "'" || (!quote && !open.length)) return null
  const tail = (quote ? '"' : '') + open.map((o) => (o === 'array' ? '}' : ')')).reverse().join('')
  const silent = !quote && !shut && open.length === 1 && open[0] === 'call' && BEGINS_WITH_CALL.test(text.slice(1))
  // an array constant left open Excel corrects only alone, in a formula begun by an array constant
  const offered = !silent && (!open.includes('array') || (open.length === 1 && /^\s*\{/.test(text.slice(1))))
  // Excel's correction drops the spaces after the formula's first character
  return { closed: text + tail, correction: text[0] + (text.slice(1) + tail).replace(/^\s+/, ''), silent, offered, calls }
}
