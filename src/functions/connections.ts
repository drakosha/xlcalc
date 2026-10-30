/**
 * The functions that reach outside the workbook: an OLAP cube's connection (the cube functions), a real-time data
 * server (RTD), the web (WEBSERVICE, IMAGE), Microsoft's services (STOCKHISTORY, TRANSLATE, DETECTLANGUAGE) and add-ins
 * (CALL, REGISTER.ID, EUROCONVERT). Without what they reach they answer as Excel for Windows answers where it reaches
 * nothing (the set `connections-held`): the cube functions and RTD #N/A whatever their arguments, WEBSERVICE #VALUE!,
 * IMAGE and STOCKHISTORY #BLOCKED! once their arguments pass, EUROCONVERT without its add-in #NAME?.
 */
import type { FnDef } from '../args.js'
import { E, type CErr, type V } from '../values.js'

/**
 * A function of connections: its cell's place first (`cell`), which a host's answer takes (`outside.ts`), and an array
 * it may answer; `fn` answers without the host.
 */
const outward = (args: string, fn: (...a: any[]) => V, more: Partial<FnDef> = {}): FnDef => ({ args, cell: true, array: true, fn: (_s: number, _r: number, _c: number, ...a: unknown[]) => fn(...a), ...more })

/** A function whose every call answers `e` without the host, its arguments taken as they are (an error among them too). */
const answering = (args: string, e: CErr, rep?: number): FnDef => outward(args, () => e, rep ? { rep } : {})

/** IMAGE's source is a web address; sizing 0 to 3 (fit, fill, original, custom). */
const WEB = /^https?:\/\//i

export const functions: Record<string, FnDef> = {
  CUBEKPIMEMBER: answering('a a a a?', E.NA),
  CUBEMEMBER: answering('a a a?', E.NA),
  CUBEMEMBERPROPERTY: answering('a a a', E.NA),
  CUBERANKEDMEMBER: answering('a a a a?', E.NA),
  CUBESET: answering('a a a? a? a?', E.NA),
  CUBESETCOUNT: answering('a', E.NA),
  CUBEVALUE: answering('a a?', E.NA, 1),
  RTD: answering('a a a', E.NA, 1),
  WEBSERVICE: outward('s', () => E.VALUE),
  IMAGE: outward('s s? n? n? n?', (source: string, _alt?: string, sizing = 0) => (!WEB.test(source) || sizing < 0 || sizing > 3 ? E.VALUE : E.BLOCKED)),
  // Excel loads as many properties as a function newer than Excel 2007 takes arguments (the set arity)
  STOCKHISTORY: outward('s d d? n? n? n?', () => E.BLOCKED, { rep: 1, max: 254 }),
  // Excel 16.0.20430 for Windows has neither (#NAME?); Microsoft's service unreached, as STOCKHISTORY's
  TRANSLATE: outward('s s? s?', () => E.BLOCKED),
  DETECTLANGUAGE: outward('s', () => E.BLOCKED),
  // Excel opens no file with CALL or REGISTER.ID in a formula
  CALL: answering('a', E.VALUE, 1),
  'REGISTER.ID': answering('a a a?', E.VALUE),
  EUROCONVERT: answering('a a a a? a?', E.NAME),
}

/** The functions whose formula keeps the value its file stores until it is calculated (`held.ts`): those above. */
export const STORED_FUNCTIONS: ReadonlySet<string> = new Set(Object.keys(functions))
