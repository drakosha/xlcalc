/**
 * The lexer of formula text for translation between locales. It does not parse: it only finds the pieces a
 * locale changes (numbers, function names, TRUE and FALSE, error literals, separators, those of array
 * constants too) and copies the rest (strings, sheet names, references, operators, whitespace) as raw pieces,
 * so text the engine cannot parse translates as well. Every character of the text belongs to exactly one piece.
 */
import { caseKey } from '../collate.js'
import { BOOK_SHEET, bracketEnd, callFollows, closingQuote, IDENT_RULE, numberSource } from '../syntax/scan.js'
import type { Message } from './messages.js'

/** Separators of formula text on one side of a translation. */
export interface Syntax {
  decimal: string
  list: string
  arrayColumn: string
  arrayRow: string
}

/** The stored (invariant) syntax: Excel's file format. */
export const INVARIANT: Syntax = { decimal: '.', list: ',', arrayColumn: ',', arrayRow: ';' }

/** Error literals of Excel, invariant spelling. */
export const ERROR_LITERALS = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#BUSY!', '#UNKNOWN!', '#PYTHON!']

/** `tbl`: the brackets of a structured reference (`[[#Data],[Qty]]`), whose specifiers and separators a locale changes. */
export type PieceKind = 'num' | 'fn' | 'word' | 'err' | 'list' | 'col' | 'row' | 'raw' | 'tbl'

export interface Piece {
  k: PieceKind
  /** Position in the text. */
  at: number
  text: string
  /** Inside an array constant. */
  arr: boolean
}

export interface LexFail {
  at: number
  msg: Message
}

// the identifier rule of the engine's tokenizer
const IDENT = new RegExp(IDENT_RULE, 'uy')

const escape = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')

/** The number rule of a syntax: `1.5`, `.5`, `5.`, `1.5E-07` with its decimal separator. */
export const numberRule = (syn: Syntax) => new RegExp(numberSource(escape(syn.decimal)), 'y')

/**
 * A structured reference's brackets with its specifiers in other words (`#Data` as `#Данные`, `words` mapping upper-case
 * words of one side to the other's) and its items' separator `from` as `to`; null when the brackets do not close. A
 * column written alone (`Sales[1,5]`, `[@1,5]`) has no items: a separator there is part of its name.
 */
export function tableText(run: string, words: ReadonlyMap<string, string>, from: string, to: string): string {
  let out = ''
  let depth = 0
  const items = run.slice(1).trimStart().startsWith('[')
  for (let j = 0; j < run.length; j++) {
    const ch = run[j]
    if (ch === "'") {
      out += run.slice(j, j + 2)
      j++
      continue
    }
    if (ch === '[') depth++
    else if (ch === ']') depth--
    else if (ch === from && depth === 1 && items) {
      out += to
      continue
    } else if (ch === '#' && run[j - 1] === '[') {
      const end = run.indexOf(']', j)
      const w = run.slice(j, end < 0 ? run.length : end)
      out += words.get(caseKey(w)) ?? w
      j += w.length - 1
      continue
    }
    out += ch
  }
  return out
}

/**
 * The pieces of `src` read with `syn`; `errors` are the error literals of this side (matched in any case,
 * the longest first).
 */
export function lex(src: string, syn: Syntax, errors: readonly string[], num: RegExp): Piece[] | LexFail {
  const out: Piece[] = []
  let depth = 0
  let i = 0
  const push = (k: PieceKind, len: number) => {
    out.push({ k, at: i, text: src.slice(i, i + len), arr: depth > 0 })
    i += len
  }
  while (i < src.length) {
    const ch = src[i]
    if (ch === '"' || ch === "'") {
      const j = closingQuote(src, i)
      if (j < 0) return { at: i, msg: { key: ch === '"' ? 'unterminatedString' : 'unterminatedSheet' } }
      // a quoted sheet name keeps its `!`
      push('raw', j + 1 - i + (ch === "'" && src[j + 1] === '!' ? 1 : 0))
      continue
    }
    if (ch === '#') {
      const head = caseKey(src.slice(i))
      const e = errors.find((x) => head.startsWith(caseKey(x)))
      if (e) {
        push('err', e.length)
        continue
      }
    }
    // the brackets of a structured reference; a reference to another workbook (`[1]Sheet!A1`) is copied as it is
    if (ch === '[' && depth === 0) {
      const end = bracketEnd(src, i)
      push(BOOK_SHEET.test(src.slice(end)) ? 'raw' : 'tbl', end - i)
      continue
    }
    if (ch === '{') depth++
    if (ch === '}' && depth > 0) {
      depth--
      out.push({ k: 'raw', at: i++, text: ch, arr: true })
      continue
    }
    if (depth > 0 ? ch === syn.arrayColumn || ch === syn.arrayRow : ch === syn.list) {
      push(depth === 0 ? 'list' : ch === syn.arrayColumn ? 'col' : 'row', 1)
      continue
    }
    num.lastIndex = i
    const n = num.exec(src)
    if (n) {
      push('num', n[0].length)
      continue
    }
    IDENT.lastIndex = i
    const id = IDENT.exec(src)
    if (id) {
      // in an array constant a separator ends a word: `\` may be one (`{TRUE\FALSE}`), and a name holds no `\`
      if (depth > 0) id[0] = id[0].split(syn.arrayColumn)[0].split(syn.arrayRow)[0]
      const end = i + id[0].length
      // an unquoted sheet name keeps its `!`
      if (src[end] === '!') {
        push('raw', id[0].length + 1)
        continue
      }
      push(callFollows(src, end) ? 'fn' : 'word', id[0].length)
      continue
    }
    push('raw', 1)
  }
  return out
}
