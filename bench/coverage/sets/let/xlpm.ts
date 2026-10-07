/**
 * LET's names in Excel's file spelling: an xlsx file writes a name `_xlpm.x` where LET declares it and where it
 * is used in its scope (after its own value, in the later values and the calculation), while the same word
 * elsewhere is a defined name. The engine's text has no prefix; this adds it, by the tokens of the formula.
 */
import { tokenize } from '../../../../src/syntax.js'

interface Frame {
  /** The parentheses of a LET call: its arguments are counted and its names declared. */
  let: boolean
  arg: number
  names: string[]
  /** The name whose value is being read: it is visible once the value ends. */
  pending?: string
}

export function xlpm(formula: string): string {
  const toks = tokenize(formula)
  const frames: Frame[] = []
  const at: number[] = []
  toks.forEach((t, k) => {
    const top = frames.at(-1)
    const prev = toks[k - 1]
    if (t.t === '(' || t.t === '{') frames.push({ let: t.t === '(' && prev?.t === 'fn' && prev.v === 'LET', arg: 0, names: [] })
    else if (t.t === ')' || t.t === '}') frames.pop()
    else if (t.t === ',' && top?.let) {
      if (top.arg % 2 === 1 && top.pending !== undefined) {
        top.names.push(top.pending.toUpperCase())
        top.pending = undefined
      }
      top.arg++
    } else if (t.t === 'ident') {
      const declares = top?.let && top.arg % 2 === 0 && (prev.t === '(' || prev.t === ',') && toks[k + 1].t === ','
      if (declares) top!.pending = t.v
      if (declares || frames.some((f) => f.let && f.names.includes(t.v.toUpperCase()))) at.push(t.at)
    }
  })
  let out = formula
  for (const p of at.reverse()) out = `${out.slice(0, p)}_xlpm.${out.slice(p)}`
  return out
}
