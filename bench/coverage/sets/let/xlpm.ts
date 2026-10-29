/**
 * LET's names and LAMBDA's parameters in Excel's file spelling: an xlsx file writes a name `_xlpm.x` where LET or
 * LAMBDA declares it and where it is used in its scope (a LET name after its own value, in the later values and the
 * calculation; a parameter in the calculation), while the same word elsewhere is a defined name; a function's name
 * as a value is `_xleta.SUM`. The engine's text has no prefix; this adds it, by the tokens of the formula.
 */
import { FUNCTION_NAMES, SYNTAX } from '../../../../src/functions.js'
import { boolWord } from '../../../../src/syntax/parse.js'
import { tokenize } from '../../../../src/syntax.js'

interface Frame {
  /** The parentheses of a LET call: its arguments are counted and its names declared. */
  let: boolean
  /** The parentheses of a LAMBDA call: its parameters are declared until an argument is not a name alone. */
  lambda?: boolean
  arg: number
  names: string[]
  /** The name whose value is being read: it is visible once the value ends. */
  pending?: string
}

export function xlpm(formula: string): string {
  const toks = tokenize(formula)
  const frames: Frame[] = []
  const at: [number, string][] = []
  toks.forEach((t, k) => {
    const top = frames.at(-1)
    const prev = toks[k - 1]
    const opens = t.t === '(' && prev?.t === 'fn'
    if (t.t === '(' || t.t === '{') frames.push({ let: opens && prev.v === 'LET', lambda: opens && prev.v === 'LAMBDA', arg: 0, names: [] })
    else if (t.t === ')' || t.t === '}') frames.pop()
    else if (t.t === ',' && top?.let) {
      if (top.arg % 2 === 1 && top.pending !== undefined) {
        top.names.push(top.pending.toUpperCase())
        top.pending = undefined
      }
      top.arg++
    } else if (t.t === ',' && top?.lambda) top.arg++
    // a LET name or a parameter called, a function's name calling the function
    else if (t.t === 'fn' && !FUNCTION_NAMES.includes(t.v) && frames.some((f) => (f.let || f.lambda) && f.names.includes(t.v))) at.push([t.at, '_xlpm.'])
    // a LAMBDA's parameter written as a cell's word (`a1`), where it declares it and in its calculation
    else if (t.t === 'ref' && t.sheet === null && t.ref.kind === 'cell' && /^[A-Za-z]+\d+$/.test(formula.slice(t.at, t.at + t.ref.text.length))) {
      const word = t.ref.text
      const param = top?.lambda && top.arg === top.names.length && (prev.t === '(' || prev.t === ',') && toks[k + 1].t === ','
      if (param) top!.names.push(word)
      if (param || frames.some((f) => f.lambda && f.names.includes(word))) at.push([t.at, '_xlpm.'])
    } else if (t.t === 'ident' && t.sheet === undefined) {
      const declares = top?.let && top.arg % 2 === 0 && (prev.t === '(' || prev.t === ',') && toks[k + 1].t === ','
      if (declares) top!.pending = t.v
      // a parameter: a name alone in each argument before the calculation
      const param = top?.lambda && top.arg === top.names.length && (prev.t === '(' || prev.t === ',') && toks[k + 1].t === ','
      if (param) top!.names.push(t.v.toUpperCase())
      if (declares || param || frames.some((f) => (f.let || f.lambda) && f.names.includes(t.v.toUpperCase()))) at.push([t.at, '_xlpm.'])
      else if (FUNCTION_NAMES.includes(t.v.toUpperCase()) && !SYNTAX.includes(t.v.toUpperCase()) && boolWord(t.v) === undefined) at.push([t.at, '_xleta.'])
    }
  })
  let out = formula
  for (const [p, prefix] of at.reverse()) out = `${out.slice(0, p)}${prefix}${out.slice(p)}`
  return out
}
