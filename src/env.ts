/**
 * Platform globals the engine uses, looked up at call time with plain-JS fallbacks, so the library runs
 * in Node, in browsers and in bare JS engines alike.
 */
import type { Clock } from './contract.js'

interface Globals {
  performance?: { now(): number }
  crypto?: { getRandomValues(array: Uint32Array): Uint32Array }
}

const g = globalThis as unknown as Globals

/** Milliseconds for timing (`ms` of a result). */
export const nowMs = (): number => (g.performance ? g.performance.now() : Date.now())

/** A random uint32 seed: Web Crypto when available, `Math.random` otherwise. */
function randomSeed(): number {
  if (g.crypto?.getRandomValues) return g.crypto.getRandomValues(new Uint32Array(1))[0]
  return Math.floor(Math.random() * 0x100000000)
}

/** The default clock: the current moment and a random seed. */
export const currentClock = (): Clock => ({ now: new Date().toISOString(), seed: randomSeed() })
