/** Error reasons: why each cell with an error has it (a formula check, a cycle, an error from another cell or its own). */
import type { CheckCode, ErrorCode } from './contract.js'
import { valueAt, type EngineState, type Formula } from './engine.js'
import { formatAddress } from './model.js'
import { ARRAY_RESULT, quoteSheet } from './syntax.js'
import { isErr } from './values.js'

const GENERIC: Record<ErrorCode, string> = {
  DIV0: 'division by zero',
  NA: 'value not found',
  NAME: 'unknown function or name',
  NUM: 'number out of range',
  REF: 'invalid reference',
  VALUE: 'wrong type of value',
  CYCLE: 'circular reference',
  SPILL: 'array result does not fit',
  CALC: 'empty array or text too long',
  ERROR: 'error',
}

export interface Issue {
  s: number
  r: number
  c: number
  error: ErrorCode
  reason: string
  check?: CheckCode
  at?: number
}

/** Cells with an error and the reason: a formula check, a cycle, an error from another cell or its own. */
export function issues(state: EngineState): Issue[] {
  const out: Issue[] = []
  for (const [k, inv] of state.invalid) {
    const [s, r, c] = k.split(':').map(Number)
    out.push({ s, r, c, error: inv.error, reason: inv.message, check: inv.code, at: inv.at })
  }
  for (const f of state.formulas) {
    if (!f) continue
    const v = valueAt(state, f.s, f.r, f.c)
    if (!isErr(v)) continue
    if (f.spill) {
      out.push({ s: f.s, r: f.r, c: f.c, error: v.code, reason: `the result is an array of ${f.spill[0]}×${f.spill[1]} values: ${ARRAY_RESULT}`, check: 'unsupported', at: 0 })
      continue
    }
    let reason = GENERIC[v.code]
    if (v.code !== 'CYCLE') {
      const from = errorSource(state, f, v.code)
      if (from) reason = `${GENERIC[v.code]}, from ${from}`
      else if (v.code === 'NA' && f.external) reason = 'value not found (or no value for an external call)'
    }
    out.push({ s: f.s, r: f.r, c: f.c, error: v.code, reason })
  }
  return out.sort((a, b) => a.s - b.s || a.r - b.r || a.c - b.c)
}

function errorSource(state: EngineState, f: Formula, code: ErrorCode): string | null {
  for (const g of f.refs) {
    const flat = state.rt.flat(g)
    const b = state.rt.bounds(g)
    for (let i = 0; i < flat.length; i++) {
      const v = flat[i]
      if (isErr(v) && v.code === code) return `${quoteSheet(state.sheetNames[g.s])}!${formatAddress(b.r1 + Math.floor(i / b.w), b.c1 + (i % b.w))}`
    }
  }
  return null
}
