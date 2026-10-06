/**
 * Lifting by function and argument: every whitelisted function's typical call with one of its one-value
 * arguments replaced by a range of two cells (`D!B1:B2`) or a two-value array (`{1;2}`). A lifted call
 * answers an array, which at the top of a case is `#SPILL!`; a function that refuses a range there answers
 * an error of its own (Excel 365: the Analysis ToolPak functions refuse ranges but lift arrays).
 *
 * The call comes from the coverage measurement's first typical case, else from `BASE` below; a function
 * without either, a volatile one or one whose arguments are all ranges has no cases here.
 */
import { readFileSync } from 'node:fs'
import { FUNCTION_NAMES } from '../../../../src/functions.js'
import type { Add } from './common.js'

/** Typical calls of the whitelisted functions the measurement has no typical case for. */
export const BASE: Record<string, string> = {
  'ERF.PRECISE': 'ERF.PRECISE(0.745)',
  'ERFC.PRECISE': 'ERFC.PRECISE(1)',
  DURATION: 'DURATION(DATE(2018,7,1),DATE(2048,1,1),0.08,0.09,2,1)',
  MDURATION: 'MDURATION(DATE(2018,7,1),DATE(2048,1,1),0.08,0.09,2,1)',
  PRICE: 'PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)',
  YIELD: 'YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2,0)',
  COUPDAYBS: 'COUPDAYBS(DATE(2011,1,25),DATE(2011,11,15),2,1)',
  COUPDAYSNC: 'COUPDAYSNC(DATE(2011,1,25),DATE(2011,11,15),2,1)',
  COUPNCD: 'COUPNCD(DATE(2011,1,25),DATE(2011,11,15),2,1)',
  COUPNUM: 'COUPNUM(DATE(2011,1,25),DATE(2011,11,15),2,1)',
  COUPPCD: 'COUPPCD(DATE(2011,1,25),DATE(2011,11,15),2,1)',
  AMORDEGRC: 'AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,1)',
  AMORLINC: 'AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,1)',
  ACCRINTM: 'ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0.1,1000,3)',
  INTRATE: 'INTRATE(DATE(2008,2,15),DATE(2008,5,15),1000000,1014420,2)',
  PRICEMAT: 'PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),0.061,0.061,0)',
  RECEIVED: 'RECEIVED(DATE(2008,2,15),DATE(2008,5,15),1000000,0.0575,2)',
  YIELDDISC: 'YIELDDISC(DATE(2008,2,16),DATE(2008,3,1),99.795,100,2)',
  YIELDMAT: 'YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,100.0123,0)',
  ODDFPRICE: 'ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)',
  ODDFYIELD: 'ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,84.5,100,2,0)',
  ODDLPRICE: 'ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,0)',
  ODDLYIELD: 'ODDLYIELD(DATE(2008,4,20),DATE(2008,6,15),DATE(2007,12,24),0.0375,99.875,100,2,0)',
  ADDRESS: 'ADDRESS(2,3)',
  HYPERLINK: 'HYPERLINK("http://example.com","x")',
  XLOOKUP: 'XLOOKUP(3,D!L1:L6,D!M1:M6)',
  XMATCH: 'XMATCH(3,D!L1:L6)',
  PERCENTOF: 'PERCENTOF(2,4)',
  'FORECAST.LINEAR': 'FORECAST.LINEAR(30,D!AY1:AY5,D!AZ1:AZ5)',
  CONFIDENCE: 'CONFIDENCE(0.05,2.5,50)',
  MODE: 'MODE(1,2,2)',
  PERCENTILE: 'PERCENTILE(D!B1:B8,0.3)',
  PERCENTRANK: 'PERCENTRANK(D!B1:B8,15)',
  POISSON: 'POISSON(2,5,TRUE)',
  QUARTILE: 'QUARTILE(D!B1:B8,1)',
  RANK: 'RANK(15,D!B1:B8)',
  STDEV: 'STDEV(1,2,3)',
  VAR: 'VAR(1,2,4)',
  WEIBULL: 'WEIBULL(105,20,100,TRUE)',
  FINDB: 'FINDB("b","abc")',
  LEFTB: 'LEFTB("abc",2)',
  LENB: 'LENB("abc")',
  MIDB: 'MIDB("abcdef",2,3)',
  REPLACEB: 'REPLACEB("abcdef",2,3,"x")',
  RIGHTB: 'RIGHTB("abc",2)',
  SEARCHB: 'SEARCHB("B","abc")',
  TEXTAFTER: 'TEXTAFTER("a-b-c","-")',
  TEXTBEFORE: 'TEXTBEFORE("a-b-c","-")',
  ARRAYTOTEXT: 'ARRAYTOTEXT(D!L1:N2,0)',
  VALUETOTEXT: 'VALUETOTEXT(1.5,0)',
  BAHTTEXT: 'BAHTTEXT(1234)',
}

const SKIP = new Set(['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'RANDARRAY'])

/** Top-level arguments of a call's text (commas outside parentheses, braces and strings). */
export function splitArgs(inner: string): string[] {
  const out: string[] = []
  let depth = 0
  let quoted = false
  let start = 0
  for (let i = 0; i < inner.length; i++) {
    const ch = inner[i]
    if (ch === '"') quoted = !quoted
    else if (quoted) continue
    else if (ch === '(' || ch === '{') depth++
    else if (ch === ')' || ch === '}') depth--
    else if (ch === ',' && depth === 0) {
      out.push(inner.slice(start, i))
      start = i + 1
    }
  }
  if (inner.length) out.push(inner.slice(start))
  return out
}

/** The typical call of `fn` as `[name, args]`, without prefixes. */
function baseCall(fn: string, typical: Map<string, string>): string[] | null {
  const text = (typical.get(fn) ?? BASE[fn])?.replace(/_xlfn\.(_xlws\.)?/g, '')
  if (!text?.startsWith(`${fn}(`) || !text.endsWith(')')) return null
  return splitArgs(text.slice(fn.length + 1, -1))
}

/** The typical calls' arguments of the whitelisted functions that have one, the volatile ones left out, in the whitelist's order. */
export function typicalCalls(): Map<string, string[]> {
  const cases = JSON.parse(readFileSync(new URL('../../../../test/excel/coverage/cases.json', import.meta.url), 'utf8')) as { cases: { fn: string; kind: string; formula: string }[] }
  const typical = new Map<string, string>()
  for (const c of cases.cases) if (c.kind === 'typical' && !typical.has(c.fn)) typical.set(c.fn, c.formula)
  const out = new Map<string, string[]>()
  for (const fn of FUNCTION_NAMES) {
    const args = SKIP.has(fn) ? null : baseCall(fn, typical)
    if (args) out.set(fn, args)
  }
  return out
}

/** A range of several cells as an argument: those positions take ranges, they are not lifted. */
const AREA = /^(?:'[^']+'|\w+)!\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+$/

export function lift(add: Add) {
  for (const [fn, args] of typicalCalls())
    args.forEach((a, i) => {
      if (AREA.test(a)) return
      for (const v of ['D!B1:B2', '{1;2}']) add(`lift:${fn}`, fn, `${fn}(${args.map((x, j) => (j === i ? v : x)).join(',')})`)
    })
}
