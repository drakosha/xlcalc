/**
 * A formula of a file whose newer functions are called without their prefix (`noPrefixCalls`), as a generator that takes
 * a formula verbatim writes them, with the prefixes put in as `writeXlsx` writes them: `_xlfn.XLOOKUP`,
 * `_xlfn._xlws.FILTER`, and in a LET or a LAMBDA so repaired its names as `_xlpm.x` and a LAMBDA's optional parameter
 * `[y]` as `_xlop.y`, which Excel reads no name without. The rest of the text stays as written: its references, spaces,
 * sheets and the prefixes it had. The engine reads the formula repaired as it read it before, since the reader drops the
 * prefixes again (`fromFile`); a repair that would read otherwise is not made.
 */
import { wordKey } from '../../collate.js'
import { FUNCTION_COUNTS } from '../../function-words.js'
import { FormulaError } from '../../syntax/ast.js'
import { tokenize, type Tok } from '../../syntax/tokens.js'
import { bareCalls, fromFile, type BareCall } from './from-file.js'
import { XLWS } from './prefixes.js'

export interface Repaired {
  formula: string
  /** The functions given their prefix, upper-case, in the order of their first call (`noPrefixCalls`). */
  fns: string[]
}

/** A piece of the text replaced: `len` characters at `at` (none for an insertion) by `text`. */
interface Edit {
  at: number
  len: number
  text: string
}

const LOCAL = /^_xl(?:pm|op)\./i
/** A LAMBDA's optional parameter as Excel shows it (`[y]`). */
const OPTIONAL = /^\[([\p{L}_\\][\p{L}\p{N}_.\\?]*)\]$/u

const apply = (src: string, edits: Iterable<Edit>): string => {
  let out = ''
  let at = 0
  for (const e of [...edits].sort((a, b) => a.at - b.at)) {
    out += src.slice(at, e.at) + e.text
    at = e.at + e.len
  }
  return out + src.slice(at)
}

/** The formula with the prefixes its calls lack, undefined where it lacks none or the repair cannot be made. */
export function repairPrefixes(src: string): Repaired | undefined {
  const calls = bareCalls(src)
  if (!calls.length) return undefined
  const edits = new Map<number, Edit>(calls.map((c) => [c.at, { at: c.at, len: c.text.length, text: `${XLWS.has(c.name) ? '_xlfn._xlws.' : '_xlfn.'}${c.name}` }]))
  const scoped = calls.filter((c) => c.name === 'LET' || c.name === 'LAMBDA')
  if (scoped.length) {
    let toks: Tok[]
    try {
      toks = tokenize(src)
    } catch (e) {
      if (e instanceof FormulaError) return undefined
      throw e
    }
    for (const c of scoped) if (!locals(src, toks, c, edits)) return undefined
  }
  const formula = apply(src, edits.values())
  // the engine reads it as before: only prefixes came in, and the functions' names in upper case
  const upper = apply(src, calls.map((c) => ({ at: c.at, len: c.text.length, text: c.name })))
  if (fromFile(formula).formula !== fromFile(upper).formula) return undefined
  return { formula, fns: [...new Set(calls.map((c) => c.name))] }
}

/**
 * The names of a LET or a LAMBDA called without its prefix (`c`), declared and used in its scope, given theirs
 * (`edits`): a LET's names from the argument after them, a LAMBDA's parameters in its body. False where an argument
 * that declares a name is not one name alone.
 */
function locals(src: string, toks: Tok[], c: BareCall, edits: Map<number, Edit>): boolean {
  const i = toks.findIndex((t) => t.t === 'fn' && t.at === c.at)
  if (i < 0 || toks[i + 1]?.t !== '(') return false
  // the arguments as ranges of tokens, split at the commas of the call itself
  const args: [number, number][] = []
  let depth = 0
  let braces = 0
  let from = i + 2
  let closed = false
  for (let j = i + 1; j < toks.length && !closed; j++) {
    const t = toks[j].t
    if (t === '(') depth++
    else if (t === ')' && --depth === 0) {
      args.push([from, j])
      closed = true
    } else if (t === '{') braces++
    else if (t === '}') braces--
    else if (t === ',' && depth === 1 && braces === 0) {
      args.push([from, j])
      from = j + 1
    }
  }
  if (!closed) return false
  const isLet = c.name === 'LET'
  for (let k = 0; k < args.length - 1; k++) {
    if (isLet && k % 2) continue
    const [a, b] = args[k]
    const t = toks[a]
    let name: string
    if (b - a === 1 && t.t === 'ident' && t.sheet === undefined) {
      name = t.v.replace(LOCAL, '')
      if (name === t.v) edits.set(t.at, { at: t.at, len: 0, text: '_xlpm.' })
    } else {
      const text = b > a ? src.slice(t.at, toks[b].at).trimEnd() : ''
      const m = isLet ? null : OPTIONAL.exec(text)
      if (!m) return false
      name = m[1]
      edits.set(t.at, { at: t.at, len: text.length, text: `_xlop.${name}` })
    }
    const key = wordKey(name)
    for (const [p, q] of isLet ? args.slice(k + 1) : [args[args.length - 1]])
      for (let j = p; j < q; j++) {
        const u = toks[j]
        // a name of the scope, or one called (a LAMBDA a LET named), but no function's
        const used = (u.t === 'ident' && u.sheet === undefined) || (u.t === 'fn' && !FUNCTION_COUNTS.has(u.v))
        if (used && wordKey(u.v) === key && !edits.has(u.at)) edits.set(u.at, { at: u.at, len: 0, text: '_xlpm.' })
      }
  }
  return true
}
