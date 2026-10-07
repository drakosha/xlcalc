/**
 * A formula as a file stores it, in the engine's spelling: the prefixes of newer functions and of LET's names go
 * (`_xlfn.`, `_xlfn._xlws.`, `_xlpm.`, `_xludf.`), `_xlfn.ANCHORARRAY(E2)` becomes the spill reference `E2#` and
 * `_xlfn.SINGLE(x)` the operator `@x`. What the engine has no spelling for is reported: user-defined functions,
 * structured references to tables, references to other workbooks, a LET name that reads as a reference or a boolean
 * once its prefix goes. Text in quotes is left as it is.
 */
import { FormulaError } from '../../syntax/ast.js'
import { parse } from '../../syntax/parse.js'

export type Found = 'userFunction' | 'table' | 'externalLink' | 'letName'

export interface Converted {
  formula: string
  found: Found[]
}

const PREFIX = /^_(?:xlfn|xlws|xlpm|xludf)\./i
const IDENT = /[\p{L}\p{N}_.\\]/u
/** A cell or a lost reference, with its sheet or without: what `#` may follow. */
const ANCHOR = /^(?:(?:'(?:[^']|'')+'|[^\s'!(),]+)!)?(?:\$?[A-Za-z]{1,3}\$?\d+|#REF!)$/
const LET_NAME_CLASH = /^(?:[A-Za-z]{1,3}\d+|[Rr]\d*[Cc]\d*|TRUE|FALSE)$/

/** The end (past the closing quote) of a quoted run starting at `i`: `""` and `''` inside are kept. */
function quoted(src: string, i: number): number {
  const q = src[i]
  let j = i + 1
  while (j < src.length) {
    if (src[j] === q) {
      if (src[j + 1] !== q) return j + 1
      j++
    }
    j++
  }
  return src.length
}

/** The index of the parenthesis closing the one at `open`, outside quotes; -1 when there is none. */
function closing(src: string, open: number): number {
  let depth = 0
  for (let i = open; i < src.length; i++) {
    const c = src[i]
    if (c === '"' || c === "'") i = quoted(src, i) - 1
    else if (c === '(') depth++
    else if (c === ')' && --depth === 0) return i
  }
  return -1
}

/** The end of a bracketed run (`[1]`, `Table1[[#This Row],[Col]]`) starting at `i`. */
function bracket(src: string, i: number): number {
  let depth = 0
  for (let j = i; j < src.length; j++) {
    if (src[j] === "'" && src[j + 1] !== undefined) j++
    else if (src[j] === '[') depth++
    else if (src[j] === ']' && --depth === 0) return j + 1
  }
  return src.length
}

/** `@x` where `@` takes the whole of `x` (a reference, a name, a call, a constant), else `@(x)`. */
function single(x: string): string {
  try {
    if (parse(`@${x}`).k === 'at') return `@${x}`
  } catch (e) {
    if (!(e instanceof FormulaError)) throw e
  }
  return `@(${x})`
}

export function fromFile(src: string): Converted {
  if (!/[_[:]/.test(src)) return { formula: src, found: [] }
  const found = new Set<Found>()
  let out = ''
  let i = 0
  while (i < src.length) {
    const c = src[i]
    if (c === '"' || c === "'") {
      const end = quoted(src, i)
      if (c === "'" && src[i + 1] === '[' && src[end] === '!') found.add('externalLink')
      out += src.slice(i, end)
      i = end
      continue
    }
    if (c === '[') {
      const end = bracket(src, i)
      found.add(IDENT.test(src[i - 1] ?? '') || /[@#[]/.test(src[i + 1] ?? '') ? 'table' : 'externalLink')
      out += src.slice(i, end)
      i = end
      continue
    }
    if (c !== '_' || IDENT.test(src[i - 1] ?? '') || !PREFIX.test(src.slice(i, i + 7))) {
      out += c
      i++
      continue
    }
    // one prefix or a chain of them (`_xlfn._xlws.`) before a function or a LET name
    let kind = ''
    let m: RegExpExecArray | null
    while ((m = PREFIX.exec(src.slice(i, i + 7)))) {
      kind = m[0].toLowerCase()
      i += m[0].length
    }
    const name = /^[\p{L}_\\][\p{L}\p{N}_.]*/u.exec(src.slice(i))?.[0] ?? ''
    const upper = name.toUpperCase()
    if (kind === '_xludf.') found.add('userFunction')
    if (kind === '_xlpm.' && LET_NAME_CLASH.test(name)) found.add('letName')
    const open = i + name.length
    const close = src[open] === '(' && (upper === 'SINGLE' || upper === 'ANCHORARRAY') ? closing(src, open) : -1
    if (close >= 0) {
      const inner = fromFile(src.slice(open + 1, close).trim())
      inner.found.forEach((x) => found.add(x))
      if (upper === 'SINGLE' || ANCHOR.test(inner.formula)) {
        out += upper === 'SINGLE' ? single(inner.formula) : `${inner.formula}#`
        i = close + 1
        continue
      }
    }
    out += name
    i += name.length
  }
  return { formula: out, found: [...found] }
}
